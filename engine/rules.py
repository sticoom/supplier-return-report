"""纯函数计算规则：原因筛选、透视、系数、分摊、取价、取整。无 IO。"""
from __future__ import annotations

import calendar
from decimal import ROUND_HALF_UP, Decimal

from engine.models import (AgreementInfo, DlmAgg, DlmRow, InboundRow,
                           InspectionBatch, ReturnOrderRow)

QUALITY_REASONS = ("DEFECTIVE", "MISSING_PARTS", "QUALITY_UNACCEPTABLE")
KNOWN_VERSIONS = {"V2版本", "V3版", "V4版"}

# 阈值表：从高到低，pass_rate >= threshold 即命中；1.0（100%）单独成支
_V3_TABLE = ((1.0, 0.0), (0.95, 0.2), (0.90, 0.5), (0.80, 0.8))
_V4_TABLE = ((1.0, 0.0), (0.97, 0.2), (0.95, 0.5), (0.90, 0.8))
_EPS = 1e-9  # 浮点安全边界（如 19/20 与 0.95 的表示误差）


def normalize_reason(reason_raw: str, fbm: bool = False) -> str | None:
    s = (reason_raw or "").strip()
    if not s:
        return None
    if fbm:
        if not s.upper().startswith("CR-"):
            return None
        s = s[3:]
    code = s.split("(", 1)[0].strip().upper()
    return code if code in QUALITY_REASONS else None


def filter_quality(rows: list[ReturnOrderRow], fbm: bool = False) -> list[ReturnOrderRow]:
    return [r for r in rows if normalize_reason(r.reason_raw, fbm=fbm)]


def pivot_quality(fba_rows, fbm_rows) -> dict[str, dict[str, int]]:
    pivot: dict[str, dict[str, int]] = {}
    for row, fbm in [(r, False) for r in fba_rows] + [(r, True) for r in fbm_rows]:
        code = normalize_reason(row.reason_raw, fbm=fbm)
        if code is None or not row.sku:
            continue
        bucket = pivot.setdefault(row.sku, dict.fromkeys(QUALITY_REASONS, 0))
        bucket[code] += row.qty
    return pivot


def aggregate_dlm(rows: list[DlmRow]) -> dict[str, DlmAgg]:
    agg: dict[str, DlmAgg] = {}
    for r in rows:
        if not r.sku:
            continue
        cur = agg.get(r.sku)
        if cur is None:
            agg[r.sku] = DlmAgg(r.sku, r.default_supplier, r.other_supplier,
                                r.sales_qty, r.return_qty)
        else:  # 同 SKU 多行（多店铺/MSKU）：数量求和，供应商取 first
            cur.sales_qty += r.sales_qty
            cur.return_qty += r.return_qty
    return agg


def batch_pass_rate(batches: list[InspectionBatch], supplier: str, month: str) -> float | None:
    s = supplier.strip()
    total = ok = 0
    for b in batches:
        if b.supplier.strip() == s and b.month == month:
            total += b.count
            if b.result.strip() == "合格":
                ok += b.count
    return ok / total if total else None


def batch_matrix(batches: list[InspectionBatch]) -> dict[str, dict[str, tuple[int, int]]]:
    """供应商×月份 批次矩阵：{供应商全名: {'YYYY-MM': (总批数, 不合格批数)}}。
    第 4 节「供应商批次合格率」sheet 数据源，口径与 wiki《2026年验货数据报表》
    「供应商」sheet 一致（合格率 = 合格批数÷总批数，非「合格」即计不合格批数）。"""
    out: dict[str, dict[str, tuple[int, int]]] = {}
    for b in batches:
        cells = out.setdefault(b.supplier.strip(), {})
        total, failed = cells.get(b.month, (0, 0))
        cells[b.month] = (total + b.count,
                          failed + (0 if b.result.strip() == "合格" else b.count))
    return out


def resolve_supplier_aliases(batches: list[InspectionBatch], full_names, prefer_sets) -> tuple:
    """汇总页的供应商是简称（如 云晴/蓓圣美）→ 映射为全名。

    prefer_sets：按业务优先级排列的候选全名集合列表（如 [入库供应商, DLM供应商, 协议表]），
    逐层找「包含简称的唯一全名」——先用最窄的层消歧（同名并存时，真正在交货的
    主体优先，如 东莞市云晴云佑电子 vs 已被替换的 深圳市云晴云佑科技）。
    全部层级都无唯一匹配 → 未解决。返回 (新列表, 未解决简称列表)。
    """
    fulls = [f.strip() for f in set(list(full_names) + [x for ps in prefer_sets for x in ps])
             if f and f.strip()]
    tiers = [[f.strip() for f in set(ps) if f and f.strip()] for ps in prefer_sets]
    resolved, pending = {}, []
    out = []
    for b in batches:
        s = b.supplier.strip()
        if s in fulls:
            out.append(b)
            continue
        if s not in resolved and s not in pending:
            hit = ""
            for tier in tiers:
                cands = [f for f in tier if s in f]
                if len(cands) == 1:
                    hit = cands[0]
                    break
            if hit:
                resolved[s] = hit
            else:
                pending.append(s)
        if s in resolved:
            out.append(InspectionBatch(resolved[s], b.month, b.result, b.count))
        else:
            out.append(b)
    return out, pending


def lookup_agreement(agreements, supplier: str) -> AgreementInfo | None:
    s = supplier.strip()
    for a in agreements:
        if a.supplier.strip() == s:
            return a
    return None


def agreement_label(agreements, supplier: str) -> str:
    info = lookup_agreement(agreements, supplier)
    if info is None:
        return "未匹配协议"
    if info.version in KNOWN_VERSIONS:        # 脏值（如误填"是"）不当作版本
        return info.version
    return "是(版本未知)" if info.signed else "否"


def coefficient(version: str, pass_rate: float | None) -> float:
    if pass_rate is None:                     # 规则19：本月无验货 → 0（优先级最高）
        return 0.0
    if version not in ("V3版", "V4版"):       # 规则16：未签/V2/版本异常 → 1.0
        return 1.0
    table = _V3_TABLE if version == "V3版" else _V4_TABLE
    for threshold, coef in table:
        if pass_rate >= threshold - _EPS:
            return coef
    return 1.2


def month_end(report_month: str) -> str:
    y, m = map(int, report_month.split("-"))
    return f"{y:04d}-{m:02d}-{calendar.monthrange(y, m)[1]:02d}"


def quarter_of(month: str) -> str:
    y, m = month.split("-")
    return f"{y}-Q{(int(m) - 1) // 3 + 1}"


def quarter_months(quarter: str) -> list[str]:
    y, q = quarter.split("-Q")
    start = int(y) * 12 + (int(q) - 1) * 3
    return [f"{i // 12:04d}-{i % 12 + 1:02d}" for i in range(start, start + 3)]


def first_price(inbounds: list[InboundRow], sku: str, supplier: str,
                report_month: str) -> float | None:
    """该 SKU×该供应商 在报告月月末及之前**首次入库**的单价（含税）。

    人工结果实证（2026-07 天鑫 13 SKU：12 家吻合 + 1 家人工笔误）：
    人工取的是年度首次入库价，而非最近一次——调价后仍按老价考核。
    同日多条取最早出现的记录。
    """
    cutoff = month_end(report_month)
    best: float | None = None
    best_key: tuple | None = None
    for i, r in enumerate(inbounds):
        if r.sku == sku and r.supplier == supplier and r.date <= cutoff:
            key = (r.date, i)                 # 同日取先出现的记录
            if best_key is None or key < best_key:
                best_key, best = key, r.unit_price
    return best


def delivery_shares(inbounds, sku: str, report_month: str) -> dict[str, float]:
    """SKU 各供应商交货数量占比（上传采购入库单中入库日期≤报告月月末的全部交货数据）。

    不区分一供/二供：谁交过货谁按占比分担该 SKU 的质量退货费用。无交货返回 {}。
    """
    cutoff = month_end(report_month)
    qty: dict[str, float] = {}
    for r in inbounds:
        if r.sku == sku and r.date and r.date <= cutoff:
            qty[r.supplier] = qty.get(r.supplier, 0.0) + r.qty
    total = sum(qty.values())
    if total <= 0:
        return {}
    return {s: v / total for s, v in sorted(qty.items(), key=lambda kv: -kv[1]) if v > 0}


def quality_return_rate(quality_qty: float, sales_qty: float) -> float | None:
    return quality_qty / sales_qty if sales_qty else None


def round2(x: float) -> float:
    return float(Decimal(str(x)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def round_undertaken(deduction: float, coefficient: float) -> int:
    return int(Decimal(str(deduction * coefficient)).quantize(
        Decimal("1"), rounding=ROUND_HALF_UP))
