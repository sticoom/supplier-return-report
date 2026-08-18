"""输入文件解析：4 个月度文件 + 2 种参考文件 → 类型化行对象。只做解析，不做业务规则。"""
from __future__ import annotations

import csv
import io
import re
from datetime import datetime
from pathlib import Path

import pandas as pd

from engine.models import (AgreementInfo, DlmRow, InboundRow, InspectionBatch,
                           ReturnOrderRow)

KNOWN_VERSIONS = {"V2版本", "V3版", "V4版"}
_ROW_PREFIX = re.compile(r"(?m)^\[row=\d+\]\s*")


def _norm(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip()


def _num(v, default: float = 0.0) -> float:
    s = _norm(v)
    if not s or s in {"nan", "None", "/"}:
        return default
    return float(str(s).replace(",", ""))


def _iso_date(v) -> str:
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    s = _norm(v)
    if not s:
        return ""
    ts = pd.to_datetime(s, errors="coerce")   # 「合计」总计行等脏日期 → 置空，由调用方跳过
    return "" if pd.isna(ts) else ts.strftime("%Y-%m-%d")


def _norm_month(v) -> str:
    if isinstance(v, datetime):
        return v.strftime("%Y-%m")
    s = _norm(v)
    y, _, m = s.partition("-")
    if y and m:
        return f"{int(y):04d}-{int(m):02d}"
    return s


def _rows(path, sheet=0) -> list[dict]:
    df = pd.read_excel(path, sheet_name=sheet, dtype=object)
    return df.to_dict("records")


def _ret_order(rec, sku_col, comment_col) -> ReturnOrderRow:
    return ReturnOrderRow(
        order_id=_norm(rec.get("订单号")), sku=_norm(rec.get(sku_col)),
        qty=int(_num(rec.get("退货数量"))), reason_raw=_norm(rec.get("退货原因")),
        return_time=_norm(rec.get("退货时间")), buyer_comment=_norm(rec.get(comment_col)),
    )


def load_fba(path) -> list[ReturnOrderRow]:
    rows = [_ret_order(r, "sku", "买家备注") for r in _rows(path)]
    return [r for r in rows if r.order_id or r.sku]


def load_fbm(path) -> list[ReturnOrderRow]:
    rows = [_ret_order(r, "SKU", "备注") for r in _rows(path)]
    return [r for r in rows if r.order_id or r.sku]


def _flatten_two_level(cols) -> list[str]:
    out = []
    for a, b in cols:
        a, b = _norm(a), _norm(b)
        out.append(a if (not b or b.startswith("Unnamed") or b == "nan") else f"{a}_{b}")
    return out


def load_dlm(path) -> list[DlmRow]:
    df = pd.read_excel(path, sheet_name=0, header=[0, 1], dtype=object)
    df.columns = _flatten_two_level(df.columns)
    rows = []
    for rec in df.to_dict("records"):
        sku = _norm(rec.get("SKU"))
        if not sku:
            continue
        rows.append(DlmRow(
            sku=sku,
            default_supplier=_norm(rec.get("默认供应商")),
            other_supplier=_norm(rec.get("其他供应商")),
            sales_qty=_num(rec.get("销量")), return_qty=_num(rec.get("退货量")),
        ))
    return rows


def load_inbound(path) -> list[InboundRow]:
    rows = []
    for rec in _rows(path):
        sku, supplier, d = _norm(rec.get("物料编码")), _norm(rec.get("供应商")), _iso_date(rec.get("入库日期"))
        if not (sku and supplier and d):
            continue
        rows.append(InboundRow(date=d, supplier=supplier, sku=sku,
                               qty=_num(rec.get("实收数量")), unit_price=_num(rec.get("单价"))))
    return rows


def load_inspection(path, sheet_name: str = "26年验货原始数据") -> list[InspectionBatch]:
    out = []
    for rec in _rows(path, sheet=sheet_name):
        sup, month = _norm(rec.get("供应商")), _norm_month(rec.get("月份"))
        if not sup or not month:
            continue
        out.append(InspectionBatch(supplier=sup, month=month, result=_norm(rec.get("质检结果"))))
    return out


def load_agreements(path) -> list[AgreementInfo]:
    if str(path).lower().endswith(".csv"):
        return _load_agreements_csv(path)
    out = []
    for rec in _rows(path):
        sup = _norm(rec.get("供应商名称"))
        if not sup:
            continue
        version = _norm(rec.get("质量协议版本"))
        out.append(AgreementInfo(sup, _norm(rec.get("质量协议")) == "是",
                                 version if version in KNOWN_VERSIONS else ""))
    return out


def _load_agreements_csv(path) -> list[AgreementInfo]:
    text = _ROW_PREFIX.sub("", Path(path).read_text(encoding="utf-8-sig"))
    records = [r for r in csv.reader(io.StringIO(text)) if r and any(c.strip() for c in r)]
    if not records:
        return []
    header = [c.strip() for c in records[0]]
    name_i = header.index("供应商名称")
    signed_i = header.index("质量协议")
    version_i = signed_i + 1        # 版本列表头为空，紧随质量协议列

    def cell(row, i):
        return row[i].strip() if i < len(row) else ""

    out = []
    for row in records[1:]:
        sup = cell(row, name_i)
        if not sup:
            continue
        version = cell(row, version_i)
        out.append(AgreementInfo(sup, cell(row, signed_i) == "是",
                                 version if version in KNOWN_VERSIONS else ""))
    return out
