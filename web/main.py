"""FastAPI 应用工厂：上传（按文件名识别类型、流式落盘）→ 参考库 → 生成 → 下载/历史/季度。"""
from __future__ import annotations

import re
import sqlite3
from dataclasses import asdict
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from engine import loaders
from engine.pipeline import Store, run_month

MONTH_RE = re.compile(r"^\d{4}-(0[1-9]|1[0-2])$")
CHUNK = 1 << 20  # 1MB 流式读写，支持大文件上传


def detect_kind(filename: str) -> str | None:
    """按真实导出文件名识别月度文件类型。"""
    lower = filename.lower()
    if "fbm" in lower:
        return "fbm"
    if "fba" in lower:
        return "fba"
    if "dlm" in lower:
        return "dlm"
    if "入库" in filename or "inbound" in lower:
        return "inbound"
    return None


def detect_ref_kind(filename: str) -> str | None:
    """按文件名识别参考库文件：验货数据 / 协议签订记录。"""
    if "验货" in filename or "inspection" in filename.lower():
        return "inspection"
    if "协议" in filename or "agreement" in filename.lower():
        return "agreements"
    return None


def latest_upload_path(store: Store, uploads_dir: Path, month: str, kind: str) -> Path | None:
    """该月该类型最近一次上传的文件路径（重传以最新为准）。"""
    with sqlite3.connect(store.db_path) as con:
        row = con.execute(
            "SELECT filename FROM upload_log WHERE report_month=? AND kind=? "
            "ORDER BY uploaded_at DESC, rowid DESC LIMIT 1", (month, kind)).fetchone()
    return uploads_dir / month / row[0] if row else None


def report_glob(month: str) -> str:
    y, m = month.split("-")
    return f"{int(y)}年{int(m)}月*.xlsx"


async def save_stream(upload: UploadFile, dest: Path) -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    with dest.open("wb") as f:
        while chunk := await upload.read(CHUNK):
            f.write(chunk)


def create_app(data_dir: str | Path | None = None) -> FastAPI:
    if data_dir is None:
        data_dir = Path(__file__).resolve().parent.parent / "data"
    data = Path(data_dir)
    uploads_dir = data / "uploads"
    reports_dir = data / "reports"
    ref_dir = data / "reference"

    app = FastAPI(title="供应商质量退货金额统计")
    store = Store(str(data / "app.db"))
    app.state.data_dir = data
    app.state.store = store
    static_dir = Path(__file__).resolve().parent / "static"
    app.mount("/static", StaticFiles(directory=static_dir), name="static")

    def check_month(month: str) -> str:
        if not MONTH_RE.match(month or ""):
            raise HTTPException(400, f"报告月格式应为 YYYY-MM，收到：{month!r}")
        return month

    @app.get("/")
    def index():
        return FileResponse(static_dir / "index.html")

    @app.post("/api/upload")
    async def upload(month: str = Form(...), files: list[UploadFile] = File(...)):
        """月度文件与参考库文件统一入口：按文件名识别类型，参考库文件直接解析入库。"""
        check_month(month)
        saved = []
        ref = store.load_reference()
        ref_changed = False
        for f in files:
            kind = detect_kind(f.filename or "") or detect_ref_kind(f.filename or "")
            if kind is None:
                raise HTTPException(
                    400, f"无法识别文件类型：{f.filename}（应含 FBA/FBM/DLM/入库/验货/协议）")
            if kind in ("inspection", "agreements"):
                dest = ref_dir / Path(f.filename).name
                await save_stream(f, dest)
                if kind == "inspection":
                    ref.inspections = loaders.load_inspection(dest)
                else:
                    ref.agreements = loaders.load_agreements(dest)
                ref_changed = True
                saved.append({"kind": kind, "filename": dest.name,
                              "path": str(dest), "note": "参考库已更新"})
                continue
            dest = uploads_dir / month / Path(f.filename).name
            await save_stream(f, dest)
            store.log_upload(month, kind, dest.name)
            saved.append({"kind": kind, "filename": dest.name, "path": str(dest)})
        if ref_changed:
            store.save_reference(ref)
        return {"month": month, "saved": saved,
                "inspections": len(ref.inspections), "agreements": len(ref.agreements)}

    @app.get("/api/reference")
    def get_reference():
        ref = store.load_reference()
        return {"inspections": len(ref.inspections), "agreements": len(ref.agreements)}

    @app.post("/api/reference")
    async def update_reference(files: list[UploadFile] = File(...)):
        """上传新版验货/协议文件 → 立即生效；只替换传入的部分，不清空另一部分。"""
        ref = store.load_reference()
        updated: dict[str, str] = {}
        for f in files:
            kind = detect_ref_kind(f.filename or "")
            if kind is None:
                raise HTTPException(400, f"无法识别参考文件类型：{f.filename}（应含 验货/协议）")
            dest = ref_dir / Path(f.filename).name
            await save_stream(f, dest)
            if kind == "inspection":
                ref.inspections = loaders.load_inspection(dest)
            else:
                ref.agreements = loaders.load_agreements(dest)
            updated[kind] = dest.name
        store.save_reference(ref)
        return {"updated": updated, "inspections": len(ref.inspections),
                "agreements": len(ref.agreements)}

    @app.post("/api/generate")
    def generate(body: dict):
        month = check_month(body.get("month", ""))
        paths, missing = {}, []
        for kind in ("fba", "fbm", "dlm", "inbound"):
            p = latest_upload_path(store, uploads_dir, month, kind)
            if p is None:
                missing.append(kind)
            else:
                paths[kind] = str(p)
        if missing:
            raise HTTPException(400, detail={"missing": missing,
                                             "msg": f"报告月 {month} 缺少月度文件，请先上传"})
        ref = store.load_reference()
        if not ref.inspections or not ref.agreements:
            lack = []
            if not ref.inspections:
                lack.append("《2026年验货数据报表.xlsx》")
            if not ref.agreements:
                lack.append("《供应商框架合同、质量协议及廉洁协议签订记录表》")
            raise HTTPException(400,
                                 f"参考数据缺少：{'、'.join(lack)}，请先在「上传参考数据文件」处上传")
        summary = run_month(month, paths["fba"], paths["fbm"], paths["dlm"], paths["inbound"],
                            ref, str(reports_dir), store=store)
        return asdict(summary)

    @app.get("/api/download/{month}")
    def download(month: str):
        check_month(month)
        cands = sorted(reports_dir.glob(report_glob(month)), key=lambda p: p.stat().st_mtime)
        if not cands:
            raise HTTPException(404, f"报告月 {month} 尚未生成报告")
        latest = cands[-1]
        return FileResponse(latest, filename=latest.name)

    @app.get("/api/history")
    def history():
        with sqlite3.connect(store.db_path) as con:
            rows = con.execute(
                "SELECT report_month, COUNT(*), SUM(under_200) FROM month_supplier "
                "GROUP BY report_month ORDER BY report_month DESC").fetchall()
        return [{"month": m, "supplier_count": n, "low200_count": low or 0}
                for m, n, low in rows]

    @app.get("/api/history/{month}")
    def history_month(month: str):
        check_month(month)
        return store.month_suppliers(month)

    @app.get("/api/quarter/{month}")
    def quarter(month: str):
        check_month(month)
        return [asdict(r) for r in store.quarter_rows(month)]

    return app


app = create_app()
