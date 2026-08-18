"""T5：engine/pipeline.py —— 编排 + SQLite 历史 + CLI。黄金用例数字全部可手推。"""
import json

from fixtures_build import (make_agreements_csv, make_dlm_file, make_fba_file,
                            make_fbm_file, make_inbound_file, make_inspection_file)

from engine import pipeline
from engine.models import ReferenceData

JIA = "东莞市甲五金制品有限公司"
YI = "台州乙塑料制品有限公司"
BING = "宁波丙塑业有限公司"


def _golden_inputs(tmp):
    fba = make_fba_file(tmp / "fba.xlsx", [
        ("A1", "SKU001", 2, "DEFECTIVE(存在瑕疵)", "2026-07-05", "破了"),
        ("A2", "SKU001", 1, "MISSING_PARTS(部分零件缺失)", "2026-07-06", ""),
        ("A3", "SKU001", 3, "NOT_COMPATIBLE(不兼容)", "2026-07-07", ""),
        ("A4", "SKU002", 1, "QUALITY_UNACCEPTABLE(质量未达到期望)", "2026-07-08", "晃"),
        ("A5", "SKU003", 1, "DEFECTIVE(存在瑕疵)", "2026-07-09", ""),
    ])
    fbm = make_fbm_file(tmp / "fbm.xlsx", [
        ("B1", "SKU001", 1, "CR-DEFECTIVE(客户退货-存在瑕疵)", "2026-07-09", ""),
        ("B2", "SKU002", 1, "CR-DAMAGED_BY_CARRIER(客户退货-被承运人损坏)", "2026-07-10", ""),
    ])
    dlm = make_dlm_file(tmp / "dlm.xlsx", [
        ("SKU001", JIA, "", 100, 10),
        ("SKU002", YI, BING, 50, 5),
    ])
    inbound = make_inbound_file(tmp / "inb.xlsx", [
        ("2026-06-01", JIA, "SKU001", 10, 60.0),
        ("2026-07-10", JIA, "SKU001", 10, 70.0),
        ("2026-08-02", JIA, "SKU001", 5, 99.0),
        ("2026-05-01", YI, "SKU002", 20, 30.0),
        ("2026-06-01", BING, "SKU002", 20, 40.0),
    ])
    return fba, fbm, dlm, inbound


def _golden_ref(tmp):
    insp = make_inspection_file(tmp / "insp.xlsx",
                                [(JIA, "2026-07", "合格")] * 19 + [(JIA, "2026-07", "不合格")]
                                + [(YI, "2026-07", "合格")] * 10
                                + [(BING, "2026-07", "合格")] * 3 + [(BING, "2026-07", "不合格")] * 2)
    agree = make_agreements_csv(tmp / "agree.csv", [(JIA, "是", "V3版"), (YI, "否", "")])
    return insp, agree


def _golden_reference(tmp):
    from engine import loaders
    insp, agree = _golden_ref(tmp)
    return ReferenceData(loaders.load_inspection(insp), loaders.load_agreements(agree))


def _run(tmp_path):
    fba, fbm, dlm, inbound = _golden_inputs(tmp_path)
    ref = _golden_reference(tmp_path)
    store = pipeline.Store(str(tmp_path / "app.db"))
    summary = pipeline.run_month("2026-07", fba, fbm, dlm, inbound, ref,
                                 str(tmp_path / "reports"), store=store)
    return summary, store


def test_run_month_golden_numbers(tmp_path):
    summary, store = _run(tmp_path)
    assert summary.supplier_count == 3 and summary.low200_count == 2
    assert summary.review_count == 2 and summary.missing_price_count == 0
    assert summary.validation_count == 3   # 未匹配协议(丙) + DLM缺失SKU + 口径差异
    assert summary.file_name == "2026年7月供应商质量退货金额汇总表.xlsx"
    assert (tmp_path / "reports" / summary.file_name).exists()

    rows = {(r["supplier"]): r for r in store.month_suppliers("2026-07")}
    jia = rows[JIA]
    assert jia["deduction"] == 280.0 and jia["coefficient"] == 0.2
    assert jia["undertaken"] == 56 and jia["under_200"] == 0
    yi = rows[YI]
    assert yi["deduction"] == 15.0 and yi["undertaken"] == 15 and yi["under_200"] == 1
    bing = rows[BING]
    assert bing["deduction"] == 20.0 and bing["undertaken"] == 20
    assert bing["agreement"] == "未匹配协议"


def test_run_month_report_data_contents(tmp_path):
    fba, fbm, dlm, inbound = _golden_inputs(tmp_path)
    ref = _golden_reference(tmp_path)
    from engine import loaders
    data = pipeline.build_report_data(
        "2026-07", loaders.load_fba(fba), loaders.load_fbm(fbm),
        loaders.load_dlm(dlm), loaders.load_inbound(inbound), ref)
    # SKU001 只列质量>0；SKU002 拆两行；数量小数保留
    jia_lines = data.suppliers[0].skus
    assert jia_lines[0].sku == "SKU001"
    assert (jia_lines[0].qty_defective, jia_lines[0].qty_missing_parts) == (3.0, 1.0)
    assert jia_lines[0].rate == 0.04 and jia_lines[0].amount == 280.0
    yi_line = [s for s in data.low200 if s.supplier == YI][0].skus[0]
    assert yi_line.qty_quality_unacceptable == 0.5 and yi_line.amount == 15.0
    assert yi_line.note == "按比例分摊"
    # 复核清单
    assert {r.order_id for r in data.review} == {"A1", "A4"}
    # 校验种类
    kinds = {v.kind for v in data.validation}
    assert kinds == {"供应商未匹配协议", "DLM缺失SKU", "口径差异"}
    # 批次合格率矩阵（第 4 节「供应商批次合格率」sheet 数据源）
    assert data.batch_matrix[JIA]["2026-07"] == (20, 1)
    assert data.batch_matrix[YI]["2026-07"] == (10, 0)
    assert data.batch_matrix[BING]["2026-07"] == (5, 2)


def test_quarter_rows_detail_and_subtotal(tmp_path):
    _, store = _run(tmp_path)
    q = store.quarter_rows("2026-07")
    details = [r for r in q if not r.is_subtotal]
    subs = {r.supplier: r for r in q if r.is_subtotal}
    assert {(r.month, r.supplier) for r in details} == {("2026-07", YI), ("2026-07", BING)}
    assert subs[YI].deduction == 15.0 and subs[BING].undertaken == 20


def test_rerun_same_month_replaces_not_accumulates(tmp_path):
    _run(tmp_path)
    summary2, store2 = _run(tmp_path)     # 同库同月重跑
    q = [r for r in store2.quarter_rows("2026-07") if not r.is_subtotal]
    assert len(q) == 2                    # 仍是 2 行，没有翻倍
    assert len(store2.month_suppliers("2026-07")) == 3


def test_store_reference_roundtrip(tmp_path):
    ref_in = _golden_reference(tmp_path)
    store = pipeline.Store(str(tmp_path / "ref.db"))
    store.save_reference(ref_in)
    ref = store.load_reference()
    assert len(ref.inspections) == 35 and len(ref.agreements) == 2
    assert ref.agreements[0].supplier == JIA and ref.agreements[0].version == "V3版"
    assert ref.inspections[0].month == "2026-07"


def test_cli_golden_end_to_end(tmp_path, capsys):
    fba, fbm, dlm, inbound = _golden_inputs(tmp_path)
    insp, agree = _golden_ref(tmp_path)
    db = str(tmp_path / "cli.db")
    out = str(tmp_path / "cli_reports")
    rc = pipeline.main(["--month", "2026-07", "--fba", fba, "--fbm", fbm, "--dlm", dlm,
                        "--inbound", inbound, "--inspection", insp, "--agreements", agree,
                        "--db", db, "--out", out])
    assert rc == 0
    payload = json.loads(capsys.readouterr().out)
    assert payload["supplier_count"] == 3 and payload["low200_count"] == 2
    store = pipeline.Store(db)
    assert len(store.month_suppliers("2026-07")) == 3
    assert len(store.load_reference().inspections) == 35   # CLI 已把参考库存进 db
