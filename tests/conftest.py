"""T6 web 层测试夹具：合成 xlsx（openpyxl 造数，不提交二进制）+ 隔离的 app/data 目录。

数据行与 tests/test_pipeline.py 黄金用例完全一致，便于断言同一组手推数字。
"""
from __future__ import annotations

import pytest

from fixtures_build import (make_agreements_csv, make_dlm_file, make_fba_file,
                            make_fbm_file, make_inbound_file, make_inspection_file)

JIA = "东莞市甲五金制品有限公司"
YI = "台州乙塑料制品有限公司"
BING = "宁波丙塑业有限公司"
MONTH = "2026-07"

# 与真实导出一致的文件名形态（按文件名识别类型的依据）
FBA_NAME = "退货(FBA)订单导出-947873100663549952.xlsx"
FBM_NAME = "退货(FBM)订单导出-947873572875407360.xlsx"
DLM_NAME = "DLM退货统计SKU导出-2026-08-17.xlsx"
INBOUND_NAME = "采购入库单_202601-07.xlsx"
INSPECTION_NAME = "2026年验货数据报表.xlsx"
AGREEMENT_NAME = "供应商协议签订记录.csv"


def make_monthly_files(tmp):
    """4 个月度合成文件，返回 {文件名: 路径}。"""
    return {
        FBA_NAME: make_fba_file(tmp / FBA_NAME, [
            ("A1", "SKU001", 2, "DEFECTIVE(存在瑕疵)", "2026-07-05", "破了"),
            ("A2", "SKU001", 1, "MISSING_PARTS(部分零件缺失)", "2026-07-06", ""),
            ("A3", "SKU001", 3, "NOT_COMPATIBLE(不兼容)", "2026-07-07", ""),
            ("A4", "SKU002", 1, "QUALITY_UNACCEPTABLE(质量未达到期望)", "2026-07-08", "晃"),
            ("A5", "SKU003", 1, "DEFECTIVE(存在瑕疵)", "2026-07-09", ""),
        ]),
        FBM_NAME: make_fbm_file(tmp / FBM_NAME, [
            ("B1", "SKU001", 1, "CR-DEFECTIVE(客户退货-存在瑕疵)", "2026-07-09", ""),
            ("B2", "SKU002", 1, "CR-DAMAGED_BY_CARRIER(客户退货-被承运人损坏)", "2026-07-10", ""),
        ]),
        DLM_NAME: make_dlm_file(tmp / DLM_NAME, [
            ("SKU001", JIA, "", 100, 10),
            ("SKU002", YI, BING, 50, 5),
        ]),
        INBOUND_NAME: make_inbound_file(tmp / INBOUND_NAME, [
            ("2026-06-01", JIA, "SKU001", 10, 60.0),
            ("2026-07-10", JIA, "SKU001", 10, 70.0),
            ("2026-08-02", JIA, "SKU001", 5, 99.0),
            ("2026-05-01", YI, "SKU002", 20, 30.0),
            ("2026-06-01", BING, "SKU002", 20, 40.0),
        ]),
    }


def make_reference_files(tmp):
    """2 个参考库合成文件（验货 xlsx + 飞书风格协议 csv）。"""
    return {
        INSPECTION_NAME: make_inspection_file(
            tmp / INSPECTION_NAME,
            [(JIA, "2026-07", "合格")] * 19 + [(JIA, "2026-07", "不合格")]
            + [(YI, "2026-07", "合格")] * 10
            + [(BING, "2026-07", "合格")] * 3 + [(BING, "2026-07", "不合格")] * 2),
        AGREEMENT_NAME: make_agreements_csv(tmp / AGREEMENT_NAME,
                                            [(JIA, "是", "V3版"), (YI, "否", "")]),
    }


@pytest.fixture()
def web_app(tmp_path):
    """create_app 工厂：数据目录隔离在 tmp_path，绝不触碰真实 data/。"""
    from web.main import create_app
    return create_app(tmp_path / "data")


@pytest.fixture()
def client(web_app):
    from fastapi.testclient import TestClient
    return TestClient(web_app)


def post_files(client, url, paths, month=None):
    """把 {文件名: 路径} 以 multipart 发到 url（字段名固定，文件名放第二元）。"""
    files = [("files", (name, open(p, "rb").read())) for name, p in paths.items()]
    data = {"month": month} if month else None
    return client.post(url, data=data, files=files)
