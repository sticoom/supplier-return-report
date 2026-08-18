"""T6：web/main.py create_app 工厂 —— TestClient 走通上传→参考→生成→下载全流程。"""
from __future__ import annotations

from conftest import (BING, INBOUND_NAME, JIA, MONTH, YI, make_monthly_files,
                      make_reference_files, post_files)


def test_index_and_upload_by_filename(client, tmp_path):
    # 首页：单页应用可访问
    r = client.get("/")
    assert r.status_code == 200 and "供应商质量退货" in r.text

    # 上传 4 个月度文件：按文件名识别类型，流式落盘到 uploads/{month}/
    monthly = make_monthly_files(tmp_path / "up")
    r = post_files(client, "/api/upload", monthly, month=MONTH)
    assert r.status_code == 200
    saved = {f["kind"]: f for f in r.json()["saved"]}
    assert set(saved) == {"fba", "fbm", "dlm", "inbound"}
    assert saved["fba"]["filename"] == list(monthly)[0] or \
        saved["fba"]["filename"].startswith("退货(FBA)")

    data_dir = client.app.state.data_dir
    for name in monthly:
        assert (data_dir / "uploads" / MONTH / name).exists()

    # 无法识别的文件名 → 400
    r = client.post("/api/upload", data={"month": MONTH},
                    files=[("files", ("随便什么.txt", b"hello"))])
    assert r.status_code == 400


def test_reference_update_effective_immediately(client, tmp_path):
    # 未上传前参考库为空
    r = client.get("/api/reference")
    assert r.status_code == 200
    assert r.json() == {"inspections": 0, "agreements": 0}

    # 上传验货 + 协议 → 立即入库生效
    refs = make_reference_files(tmp_path / "ref")
    r = post_files(client, "/api/reference", refs)
    assert r.status_code == 200
    body = r.json()
    assert body["inspections"] == 35 and body["agreements"] == 2
    assert set(body["updated"]) == {"inspection", "agreements"}

    # 只再传一份新协议 → 协议被替换、验货保留（部分更新语义，与 CLI 一致）
    from fixtures_build import make_agreements_csv
    new_agree = {"新版协议.csv": make_agreements_csv(
        tmp_path / "ref2" / "新版协议.csv", [(JIA, "是", "V4版")])}
    r = post_files(client, "/api/reference", new_agree)
    assert r.status_code == 200
    body = r.json()
    assert body["inspections"] == 35 and body["agreements"] == 1

    # 无法识别的参考文件 → 400
    r = client.post("/api/reference", files=[("files", ("别的.xlsx", b"x"))])
    assert r.status_code == 400


def test_full_flow_generate_download_history_quarter(client, tmp_path):
    # 缺文件先生成 → 400 且指出缺哪些
    r = client.post("/api/generate", json={"month": MONTH})
    assert r.status_code == 400 and set(r.json()["detail"]["missing"]) == {"fba", "fbm", "dlm", "inbound"}

    # 上传→参考→生成 全流程
    assert post_files(client, "/api/upload", make_monthly_files(tmp_path / "up"), month=MONTH).status_code == 200
    assert post_files(client, "/api/reference", make_reference_files(tmp_path / "ref")).status_code == 200

    r = client.post("/api/generate", json={"month": MONTH})
    assert r.status_code == 200
    s = r.json()
    assert s["report_month"] == MONTH
    assert s["file_name"] == "2026年7月供应商质量退货金额汇总表.xlsx"
    assert s["supplier_count"] == 3 and s["low200_count"] == 2
    assert s["review_count"] == 2 and s["missing_price_count"] == 0
    assert (client.app.state.data_dir / "reports" / s["file_name"]).exists()

    # 下载：xlsx 字节流（zip 头 PK）
    r = client.get(f"/api/download/{MONTH}")
    assert r.status_code == 200
    assert r.content[:2] == b"PK" and len(r.content) > 1000
    assert "attachment" in r.headers.get("content-disposition", "")

    # 历史列表 + 月度明细（黄金数字与 T5 一致）
    r = client.get("/api/history")
    assert {h["month"] for h in r.json()} == {MONTH}
    r = client.get(f"/api/history/{MONTH}")
    rows = {row["supplier"]: row for row in r.json()}
    assert rows[JIA]["deduction"] == 280.0 and rows[JIA]["undertaken"] == 56
    assert rows[BING]["agreement"] == "未匹配协议"

    # 季度累计：明细行 + 小计行
    r = client.get(f"/api/quarter/{MONTH}")
    q = r.json()
    assert {(row["month"], row["supplier"]) for row in q if not row["is_subtotal"]} \
        == {(MONTH, YI), (MONTH, BING)}
    subs = {row["supplier"]: row for row in q if row["is_subtotal"]}
    assert subs[YI]["deduction"] == 15.0 and subs[BING]["undertaken"] == 20

    # 未生成的月下载 → 404
    assert client.get("/api/download/2026-08").status_code == 404
