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
    """验货数据两种格式都支持：
    ① 原始明细页「26年验货原始数据」（供应商全名 + 月份 + 质检结果，每行 1 批）
    ② 「供应商」汇总页（表头含 供应商名称/统计项，行为 总批数/不合格批数/批次合格率，
       列为 1月..12月，供应商为简称）——来自 wiki 文件《2026年验货数据报表》的供应商 sheet。
    """
    out = []
    try:
        for rec in _rows(path, sheet=sheet_name):
            sup, month = _norm(rec.get("供应商")), _norm_month(rec.get("月份"))
            if not sup or not month:
                continue
            out.append(InspectionBatch(supplier=sup, month=month,
                                       result=_norm(rec.get("质检结果"))))
    except ValueError:                    # 指定 sheet 不存在 → 尝试汇总页
        out = []
    return out or _load_inspection_summary(path)


def _load_inspection_summary(path) -> list[InspectionBatch]:
    """解析「供应商」汇总页：总批数/不合格批数 × 月列 → 合成带 count 的批次记录。"""
    import re as _re
    xls = pd.ExcelFile(path)
    for sn in xls.sheet_names:
        df = pd.read_excel(path, sheet_name=sn, header=None, dtype=object)
        if df.empty:
            continue
        hdr_i = hdr = None
        for i in range(min(10, len(df))):
            vals = [str(v).strip() for v in df.iloc[i].tolist()]
            if "供应商名称" in vals and "统计项" in vals:
                hdr_i, hdr = i, vals
                break
        if hdr_i is None:
            continue
        year = datetime.now().year
        for i in range(hdr_i):            # 标题行里找年份，如「2026年供应商批次合格率统计表」
            m = _re.match(r"(\d{4})年", str(df.iloc[i, 0]).strip())
            if m:
                year = int(m.group(1))
                break
        month_cols = {}
        for j, v in enumerate(hdr):
            mm = _re.fullmatch(r"(\d{1,2})月", v)
            if mm:
                month_cols[j] = f"{year:04d}-{int(mm.group(1)):02d}"
        name_col, item_col = hdr.index("供应商名称"), hdr.index("统计项")
        out: list[InspectionBatch] = []
        cur_sup, pending = "", {}

        def _flush():
            for mon, (tot, fail) in pending.items():
                fail = int(fail or 0)
                ok = max(int(tot or 0) - fail, 0)
                if ok:
                    out.append(InspectionBatch(cur_sup, mon, "合格", ok))
                if fail:
                    out.append(InspectionBatch(cur_sup, mon, "不合格", fail))

        for i in range(hdr_i + 1, len(df)):
            sup = _norm(df.iloc[i, name_col])
            if sup and sup != cur_sup:
                _flush()
                cur_sup, pending = sup, {}
            if not cur_sup:
                continue
            item = _norm(df.iloc[i, item_col])
            if item not in ("总批数", "不合格批数"):
                continue
            idx = 0 if item == "总批数" else 1
            for j, mon in month_cols.items():
                v = _num(df.iloc[i, j])
                if v:
                    pending.setdefault(mon, [0, 0])[idx] = int(v)
        _flush()
        if out:
            return out
    return []


def load_agreements(path) -> list[AgreementInfo]:
    if str(path).lower().endswith(".csv"):
        return _load_agreements_csv(path)
    out: list[AgreementInfo] = []
    # 表头不一定是第 1 行（wiki 导出版第 3 行才是 表序/代码/名称…）→ 扫描定位
    for rec in _rows_with_header_scan(path, must_have=("供应商名称", "质量协议")):
        sup = _norm(rec.get("供应商名称"))
        if not sup:
            continue
        # 版本列的表头是空的（紧跟「质量协议」列）→ 取值在 KNOWN_VERSIONS 里的那列
        version = _norm(rec.get("质量协议版本"))
        if not version:
            for k, v in rec.items():
                if str(k).startswith("Unnamed:") and _norm(v) in KNOWN_VERSIONS:
                    version = _norm(v)
                    break
        out.append(AgreementInfo(sup, _norm(rec.get("质量协议")) == "是",
                                 version if version in KNOWN_VERSIONS else ""))
    return out


def _rows_with_header_scan(path, must_have, sheet=0, max_scan=10) -> list[dict]:
    df = pd.read_excel(path, sheet_name=sheet, header=None, dtype=object)
    for i in range(min(max_scan, len(df))):
        vals = [str(v).strip() for v in df.iloc[i].tolist()]
        if all(m in vals for m in must_have):
            hdr = [v if v and v != "nan" else f"Unnamed:{j}"
                   for j, v in enumerate(vals)]
            df2 = df.iloc[i + 1:].copy()
            df2.columns = hdr
            return df2.to_dict("records")
    return []


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
