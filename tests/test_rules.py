from engine import models as m
from engine import rules


def _ret(reason, sku="S1", qty=1, comment=""):
    return m.ReturnOrderRow("o", sku, qty, reason, "2026-07-01", comment)


# ---- 3.1 三原因筛选 + CR- 前缀 ----

def test_normalize_fba_quality_reasons():
    assert rules.normalize_reason("DEFECTIVE(存在瑕疵)") == "DEFECTIVE"
    assert rules.normalize_reason("MISSING_PARTS(部分零件缺失)") == "MISSING_PARTS"
    assert rules.normalize_reason("QUALITY_UNACCEPTABLE(质量未达到期望)") == "QUALITY_UNACCEPTABLE"
    assert rules.normalize_reason("NOT_COMPATIBLE(不兼容)") is None
    assert rules.normalize_reason("") is None


def test_normalize_fbm_requires_cr_prefix():
    assert rules.normalize_reason("CR-MISSING_PARTS(客户退货-缺失零部件)", fbm=True) == "MISSING_PARTS"
    assert rules.normalize_reason("CR-DEFECTIVE(客户退货-存在瑕疵)", fbm=True) == "DEFECTIVE"
    assert rules.normalize_reason("CR-DAMAGED_BY_CARRIER(客户退货-被承运人损坏)", fbm=True) is None
    assert rules.normalize_reason("DEFECTIVE(存在瑕疵)", fbm=True) is None


def test_pivot_quality_merges_fba_and_fbm():
    fba = [_ret("DEFECTIVE(存在瑕疵)", "S1", 2), _ret("NOT_COMPATIBLE(不兼容)", "S1", 3),
           _ret("MISSING_PARTS(部分零件缺失)", "S1", 1)]
    fbm = [_ret("CR-DEFECTIVE(客户退货-存在瑕疵)", "S1", 1)]
    pivot = rules.pivot_quality(fba, fbm)
    assert pivot == {"S1": {"DEFECTIVE": 3, "MISSING_PARTS": 1, "QUALITY_UNACCEPTABLE": 0}}


def test_filter_quality_keeps_rows_with_comment():
    rows = [_ret("DEFECTIVE(存在瑕疵)", comment="碎了"), _ret("NOT_COMPATIBLE(不兼容)")]
    kept = rules.filter_quality(rows)
    assert len(kept) == 1 and kept[0].buyer_comment == "碎了"


# ---- 3.2 DLM 聚合 ----

def test_aggregate_dlm_sums_and_takes_first():
    rows = [m.DlmRow("S1", "一供", "二供", 10, 2), m.DlmRow("S1", "一供x", "", 5, 1)]
    agg = rules.aggregate_dlm(rows)["S1"]
    assert (agg.sales_qty, agg.return_qty) == (15.0, 3.0)
    assert agg.default_supplier == "一供" and agg.other_supplier == "二供"


# ---- 3.5 系数表全分支（含边界等于） ----

def test_coefficient_v3_table():
    assert rules.coefficient("V3版", 1.0) == 0.0
    assert rules.coefficient("V3版", 0.95) == 0.2
    assert rules.coefficient("V3版", 0.9661) == 0.2
    assert rules.coefficient("V3版", 0.9499) == 0.5
    assert rules.coefficient("V3版", 0.90) == 0.5
    assert rules.coefficient("V3版", 0.8999) == 0.8
    assert rules.coefficient("V3版", 0.80) == 0.8
    assert rules.coefficient("V3版", 0.7999) == 1.2


def test_coefficient_v4_table():
    assert rules.coefficient("V4版", 1.0) == 0.0
    assert rules.coefficient("V4版", 0.97) == 0.2
    assert rules.coefficient("V4版", 0.9699) == 0.5
    assert rules.coefficient("V4版", 0.95) == 0.5
    assert rules.coefficient("V4版", 0.90) == 0.8
    assert rules.coefficient("V4版", 0.8999) == 1.2


def test_coefficient_unsigned_v2_dirty_and_no_inspection():
    assert rules.coefficient("", 0.5) == 1.0        # 未签
    assert rules.coefficient("V2版本", 0.5) == 1.0  # V2
    assert rules.coefficient("乱填", 0.5) == 1.0    # 版本脏值
    assert rules.coefficient("V3版", None) == 0.0   # 本月无验货，优先级最高
    assert rules.coefficient("", None) == 0.0


def test_batch_pass_rate_and_no_records():
    batches = [m.InspectionBatch(" 甲 ", "2026-07", "合格")] * 19
    batches.append(m.InspectionBatch("甲", "2026-07", "不合格"))
    assert rules.batch_pass_rate(batches, "甲", "2026-07") == 19 / 20
    assert rules.batch_pass_rate(batches, "甲", "2026-06") is None
    assert rules.batch_pass_rate(batches, "乙", "2026-07") is None


def test_agreement_label_variants():
    ags = [m.AgreementInfo("甲", True, "V3版"), m.AgreementInfo("乙", False, ""),
           m.AgreementInfo("丙", True, ""), m.AgreementInfo("丁", True, "是")]
    assert rules.agreement_label(ags, "甲") == "V3版"
    assert rules.agreement_label(ags, "乙") == "否"
    assert rules.agreement_label(ags, "丙") == "是(版本未知)"
    assert rules.agreement_label(ags, "丁") == "是(版本未知)"   # 脏值"是"归未知
    assert rules.agreement_label(ags, "戊") == "未匹配协议"


# ---- 3.3/3.4 取价、分摊、金额 ----

def _inb(date, sup, sku, qty, price):
    return m.InboundRow(date, sup, sku, qty, price)


def test_latest_price_latest_on_or_before_month_end():
    rows = [_inb("2026-06-01", "甲", "S1", 10, 60.0), _inb("2026-07-10", "甲", "S1", 10, 70.0),
            _inb("2026-08-02", "甲", "S1", 5, 99.0), _inb("2026-07-10", "乙", "S1", 5, 88.0)]
    assert rules.latest_price(rows, "S1", "甲", "2026-07") == 70.0  # 最近入库价（7月70，8月99不算）
    assert rules.latest_price(rows, "S1", "乙", "2026-07") == 88.0  # 各供各价
    assert rules.latest_price(rows, "S1", "甲", "2026-05") is None


def test_delivery_shares_by_receipt_qty():
    rows = [_inb("2026-06-01", "乙", "S2", 30, 1.0), _inb("2026-06-01", "丙", "S2", 10, 1.0),
            _inb("2026-08-01", "丁", "S2", 60, 1.0)]   # 8月入库不计入 7 月报告
    assert rules.delivery_shares(rows, "S2", "2026-07") == {"乙": 0.75, "丙": 0.25}
    assert rules.delivery_shares(rows, "S2", "2026-05") == {}   # 无交货数据
    assert rules.delivery_shares(rows, "S9", "2026-07") == {}


def test_quality_return_rate_zero_sales_is_none():
    assert rules.quality_return_rate(3, 100) == 0.03
    assert rules.quality_return_rate(3, 0) is None


def test_round_undertaken_half_up_and_round2():
    assert rules.round_undertaken(100.5, 1.0) == 101      # 四舍五入，非银行家舍入
    assert rules.round_undertaken(691.8, 1.0) == 692
    assert rules.round_undertaken(280.0, 0.2) == 56
    assert rules.round2(2657.044999) == 2657.04
    assert rules.round2(2657.045) == 2657.05


def test_quarter_helpers():
    assert rules.month_end("2026-07") == "2026-07-31"
    assert rules.month_end("2026-02") == "2026-02-28"
    assert rules.quarter_of("2026-07") == "2026-Q3"
    assert rules.quarter_months("2026-Q1") == ["2026-01", "2026-02", "2026-03"]


def test_batch_matrix_supplier_by_month_counts():
    """第 4 节「供应商批次合格率」sheet 数据源：供应商×月份 → (总批数, 不合格批数)。"""
    batches = [m.InspectionBatch(" 甲 ", "2026-06", "合格")] * 9 + [
        m.InspectionBatch("甲", "2026-06", "不合格"),
        m.InspectionBatch("甲", "2026-07", "合格"),
        m.InspectionBatch("乙", "2026-07", "合格"),
        m.InspectionBatch("乙", "2026-07", "不合格"),
    ]
    mx = rules.batch_matrix(batches)
    assert mx["甲"]["2026-06"] == (10, 1)       # 供应商名 strip 后合并
    assert mx["甲"]["2026-07"] == (1, 0)
    assert mx["乙"]["2026-07"] == (2, 1)
    assert rules.batch_matrix([]) == {}
