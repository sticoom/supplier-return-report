from pathlib import Path

import openpyxl
import pytest
from fixtures_build import (make_agreements_csv, make_dlm_file, make_fba_file,
                            make_fbm_file, make_inbound_file, make_inspection_file)

from engine import models


def test_models_fields():
    r = models.ReturnOrderRow("o1", "SKU1", 2, "DEFECTIVE(存在瑕疵)", "2026-07-01", "")
    assert r.qty == 2
    s = models.SupplierResult("供应商A", 100.0, None, "否", 1.0, 100, True, [])
    assert s.under_200 is True
    assert models.QUALITY_REASON_CODES == ("DEFECTIVE", "MISSING_PARTS", "QUALITY_UNACCEPTABLE")


@pytest.mark.parametrize("make", [make_fba_file, make_fbm_file, make_dlm_file,
                                  make_inbound_file, make_inspection_file])
def test_xlsx_factories_roundtrip(tmp_path, make):
    p = make(tmp_path / "f.xlsx", [])
    assert Path(p).exists()
    wb = openpyxl.load_workbook(p)
    assert wb.worksheets


def test_agreements_csv_has_row_prefix(tmp_path):
    p = make_agreements_csv(tmp_path / "a.csv", [("某供应商", "是", "V3版")])
    first = Path(p).read_text(encoding="utf-8").splitlines()[0]
    assert first.startswith("[row=3] ")
