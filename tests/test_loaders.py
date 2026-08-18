from pathlib import Path

from fixtures_build import (make_agreements_csv, make_agreements_xlsx, make_dlm_file,
                            make_fba_file, make_fbm_file, make_inbound_file,
                            make_inspection_file)

from engine import loaders


def test_load_fba_maps_columns(tmp_path):
    p = make_fba_file(tmp_path / "fba.xlsx", [
        ("A1", "SKU001", 2, "DEFECTIVE(存在瑕疵)", "2026-07-05", "破了"),
        ("A2", "SKU002", 1, "NOT_COMPATIBLE(不兼容)", "2026-07-06", ""),
    ])
    rows = loaders.load_fba(p)
    assert len(rows) == 2
    assert rows[0].order_id == "A1" and rows[0].sku == "SKU001"
    assert rows[0].qty == 2 and rows[0].buyer_comment == "破了"
    assert rows[1].reason_raw == "NOT_COMPATIBLE(不兼容)"   # 非质量行也保留


def test_load_fbm_maps_columns_and_note_as_comment(tmp_path):
    p = make_fbm_file(tmp_path / "fbm.xlsx", [
        ("B1", "SKU001", 1, "CR-MISSING_PARTS(客户退货-缺失零部件)", "2026-07-09", "缺件"),
    ])
    rows = loaders.load_fbm(p)
    assert rows[0].sku == "SKU001" and rows[0].buyer_comment == "缺件"


def test_load_dlm_two_row_header_and_values(tmp_path):
    p = make_dlm_file(tmp_path / "dlm.xlsx", [
        ("SKU001", "东莞市甲五金制品有限公司", "", 100, 10),
        ("SKU001", "东莞市甲五金制品有限公司", "", 50, 5),
        ("SKU002", "台州乙塑料制品有限公司", "宁波丙塑业有限公司", 30, 3),
    ])
    rows = loaders.load_dlm(p)
    assert len(rows) == 3            # 未聚合，聚合是 rules.aggregate_dlm 的事
    assert rows[0].default_supplier == "东莞市甲五金制品有限公司"
    assert rows[2].other_supplier == "宁波丙塑业有限公司"
    assert rows[0].sales_qty == 100 and rows[2].return_qty == 3


def test_load_inbound_price_is_untaxed_and_date_normalized(tmp_path):
    p = make_inbound_file(tmp_path / "inb.xlsx", [
        ("2026/8/4", "台州欧堡塑料制品有限公司", "DLM902001", 588, 27.2743),
        ("2026-07-31", "台州欧堡塑料制品有限公司", "DLM902001", 100, 25.5),
    ])
    rows = loaders.load_inbound(p)
    assert rows[0].date == "2026-08-04"        # '2026/8/4' 归一为 ISO
    assert rows[0].unit_price == 27.2743       # 单价列（不含税），非含税单价
    assert rows[1].date == "2026-07-31" and rows[1].qty == 100


def test_load_inbound_skips_grand_total_row(tmp_path):
    """真实采购入库单末行是「合计」总计行（入库日期列=『合计』）：不得崩溃，整行跳过。"""
    p = make_inbound_file(tmp_path / "inb.xlsx", [
        ("2026-07-31", "台州欧堡塑料制品有限公司", "DLM902001", 100, 25.5),
        ("合计", "合计", "合计", 688, 0),
    ])
    rows = loaders.load_inbound(p)
    assert len(rows) == 1 and rows[0].date == "2026-07-31"
    assert rows[0].supplier == "台州欧堡塑料制品有限公司" and rows[0].qty == 100


def test_load_inspection_strips_supplier_and_month(tmp_path):
    p = make_inspection_file(tmp_path / "insp.xlsx", [
        (" 东莞市畅艺鑫五金制品有限公司 ", "2026-03", "合格"),
        ("东莞市畅艺鑫五金制品有限公司", "2026-03", "不合格"),
    ])
    rows = loaders.load_inspection(p)
    assert rows[0].supplier == "东莞市畅艺鑫五金制品有限公司"
    assert rows[0].month == "2026-03" and rows[1].result == "不合格"


def test_load_agreements_csv_cleans_row_prefix_and_blank_version_header(tmp_path):
    p = make_agreements_csv(tmp_path / "agree.csv", [
        ("深圳市云晴云佑科技贸易有限公司", "是", "V3版"),
        ("惠州市诺尔塑胶五金有限公司", "是", "V2版本"),
        ("某未签公司", "否", ""),
        ("某脏数据公司", "是", "是"),      # 版本列误填"是" → version ''
    ])
    rows = loaders.load_agreements(p)
    assert (rows[0].supplier, rows[0].signed, rows[0].version) == (
        "深圳市云晴云佑科技贸易有限公司", True, "V3版")
    assert (rows[1].version, rows[1].signed) == ("V2版本", True)
    assert (rows[2].version, rows[2].signed) == ("", False)
    assert (rows[3].version, rows[3].signed) == ("", True)


def test_load_agreements_xlsx_variant(tmp_path):
    p = make_agreements_xlsx(tmp_path / "agree.xlsx", [("甲公司", "是", "V4版")])
    rows = loaders.load_agreements(p)
    assert rows[0].version == "V4版"
