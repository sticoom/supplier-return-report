"""编排：loaders → rules → report + SQLite 存储 + CLI。"""
from __future__ import annotations

import argparse
import json
import sqlite3
from dataclasses import asdict
from datetime import datetime
from pathlib import Path

from engine import loaders, rules
from engine.models import (AgreementInfo, InspectionBatch, QuarterRow, ReportData,
                           ReferenceData, ReviewLine, RunSummary, SkuLine,
                           SupplierResult, ValidationItem)
from engine.report import write_report, write_supplier_workbook

SCHEMA = """
CREATE TABLE IF NOT EXISTS month_supplier (
  report_month TEXT NOT NULL, supplier TEXT NOT NULL, deduction REAL NOT NULL,
  pass_rate REAL, agreement TEXT NOT NULL, coefficient REAL NOT NULL,
  undertaken INTEGER NOT NULL, under_200 INTEGER NOT NULL,
  PRIMARY KEY (report_month, supplier));
CREATE TABLE IF NOT EXISTS reference_inspection (supplier TEXT, month TEXT, result TEXT, count INTEGER DEFAULT 1);
CREATE TABLE IF NOT EXISTS reference_agreement (supplier TEXT PRIMARY KEY, signed INTEGER, version TEXT);
CREATE TABLE IF NOT EXISTS upload_log (report_month TEXT, kind TEXT, filename TEXT, uploaded_at TEXT);
"""


class Store:
    """SQLite 历史库：月度结果（同月 upsert 幂等）+ 参考数据 + 上传日志。"""

    def __init__(self, db_path: str):
        self.db_path = str(db_path)
        Path(self.db_path).parent.mkdir(parents=True, exist_ok=True)
        with sqlite3.connect(self.db_path) as con:
            con.executescript(SCHEMA)

    def upsert_month(self, report_month: str, suppliers: list[SupplierResult]) -> None:
        with sqlite3.connect(self.db_path) as con:
            con.execute("DELETE FROM month_supplier WHERE report_month=?", (report_month,))
            con.executemany(
                "INSERT INTO month_supplier VALUES (?,?,?,?,?,?,?,?)",
                [(report_month, s.supplier, s.deduction, s.pass_rate, s.agreement,
                  s.coefficient, s.undertaken, int(s.under_200)) for s in suppliers])

    def month_suppliers(self, report_month: str) -> list[dict]:
        with sqlite3.connect(self.db_path) as con:
            con.row_factory = sqlite3.Row
            cur = con.execute("SELECT * FROM month_supplier WHERE report_month=?",
                              (report_month,))
            return [dict(r) for r in cur.fetchall()]

    def quarter_rows(self, report_month: str) -> list[QuarterRow]:
        rows = []
        for month in rules.quarter_months(rules.quarter_of(report_month)):
            for r in self.month_suppliers(month):
                if r["under_200"]:
                    rows.append(QuarterRow(month, r["supplier"], r["deduction"],
                                           r["undertaken"], False))
        by_sup: dict[str, list[QuarterRow]] = {}
        for q in rows:
            by_sup.setdefault(q.supplier, []).append(q)
        subs = [QuarterRow("季度小计", sup,
                           rules.round2(sum(q.deduction for q in qs)),
                           sum(q.undertaken for q in qs), True)
                for sup, qs in by_sup.items()]
        return rows + subs

    def save_reference(self, ref: ReferenceData) -> None:
        with sqlite3.connect(self.db_path) as con:
            _ensure_count_column(con)
            con.execute("DELETE FROM reference_inspection")
            con.execute("DELETE FROM reference_agreement")
            con.executemany("INSERT INTO reference_inspection VALUES (?,?,?,?)",
                            [(b.supplier, b.month, b.result, b.count) for b in ref.inspections])
            con.executemany("INSERT OR REPLACE INTO reference_agreement VALUES (?,?,?)",
                            [(a.supplier, int(a.signed), a.version) for a in ref.agreements])

    def load_reference(self) -> ReferenceData:
        with sqlite3.connect(self.db_path) as con:
            _ensure_count_column(con)
            insp = [InspectionBatch(r[0], r[1], r[2], int(r[3] or 1)) for r in
                    con.execute("SELECT supplier, month, result, count FROM reference_inspection")]
            ags = [AgreementInfo(r[0], bool(r[1]), r[2]) for r in
                   con.execute("SELECT supplier, signed, version FROM reference_agreement")]
        return ReferenceData(insp, ags)

    def log_upload(self, report_month: str, kind: str, filename: str) -> None:
        with sqlite3.connect(self.db_path) as con:
            con.execute("INSERT INTO upload_log VALUES (?,?,?,?)",
                        (report_month, kind, filename, datetime.now().isoformat()))


def _ensure_count_column(con) -> None:
    """旧库 reference_inspection 没有 count 列（汇总页导入前建的）→ 补列，默认 1。"""
    cols = {r[1] for r in con.execute("PRAGMA table_info(reference_inspection)")}
    if cols and "count" not in cols:
        con.execute("ALTER TABLE reference_inspection ADD COLUMN count INTEGER DEFAULT 1")


def build_report_data(report_month, fba_rows, fbm_rows, dlm_rows, inbound_rows,
                      ref: ReferenceData) -> ReportData:
    pivot = rules.pivot_quality(fba_rows, fbm_rows)
    dlm_agg = rules.aggregate_dlm(dlm_rows)
    validation: list[ValidationItem] = []
    lines_by_supplier: dict[str, list[SkuLine]] = {}
    dropped = 0                        # 无交货记录被剔除的老品 SKU 数（仅用于覆盖告警）

    # 验货汇总页的供应商是简称（云晴/蓓圣美…）→ 唯一包含匹配映射为全名
    # DLM「其他供应商」可能是分号拼接的多个供应商 → 拆开后再进候选池
    dlm_suppliers = {p.strip() for a in dlm_agg.values()
                     for s in (a.default_supplier, a.other_supplier) if s
                     for p in str(s).split(";") if p.strip()}
    inbound_suppliers = {r.supplier.strip() for r in inbound_rows if r.supplier.strip()}
    full_names = ({a.supplier for a in ref.agreements} | dlm_suppliers | inbound_suppliers)
    batches, unresolved = rules.resolve_supplier_aliases(
        ref.inspections, full_names,
        [inbound_suppliers, dlm_suppliers, {a.supplier for a in ref.agreements}])
    if unresolved:
        validation.append(ValidationItem(
            "验货简称未匹配",
            "验货数据中这些供应商简称无法唯一映射为全名：" + "、".join(sorted(set(unresolved)))))
    ref = ReferenceData(batches, ref.agreements)

    for sku in sorted(pivot):
        q = pivot[sku]
        total_q = sum(q.values())
        if total_q <= 0:
            continue
        agg = dlm_agg.get(sku)
        if agg is None:
            validation.append(ValidationItem("DLM缺失SKU",
                                             f"{sku} 质量退货 {total_q} 件，DLM 表无此 SKU，未计入"))
            continue
        rate = rules.quality_return_rate(total_q, agg.sales_qty)
        shares = rules.delivery_shares(inbound_rows, sku, report_month)
        if not shares:
            # 用户拍板 2026-08-20：采购入库单（1月起）中无该 SKU 交货记录 =
            # 早已不再交货的老品 → 整体剔除：不计算、不出现在任何清单/校验里
            dropped += 1
            continue
        # 规则13：按交货数量占比分摊给所有交过货的供应商，单价取最近一次入库含税价
        multi = len(shares) > 1
        import math
        plan = [(sup, {k: (math.ceil(v * share) if v * share != int(v * share) else int(v * share))
                       for k, v in q.items()},
                 rules.latest_price(inbound_rows, sku, sup, report_month),
                 "按交货比例分摊" if multi else "") for sup, share in shares.items()]
        for sup, qq, price, note in plan:
            amount = rules.round2(sum(qq.values()) * price) if price is not None else None
            if price is None:      # 规则12：缺价 → 金额留空、不计应扣、数量照常展示
                validation.append(ValidationItem(
                    "缺单价", f"{sku} @ {sup} 报告月月末前无该供应商入库单价"))
            lines_by_supplier.setdefault(sup, []).append(SkuLine(
                sku=sku, sales_qty=agg.sales_qty, return_qty=agg.return_qty,
                qty_defective=qq["DEFECTIVE"], qty_missing_parts=qq["MISSING_PARTS"],
                qty_quality_unacceptable=qq["QUALITY_UNACCEPTABLE"],
                rate=rate, unit_price=price, amount=amount, note=note))

    results = []
    for sup, lines in lines_by_supplier.items():
        deduction = rules.round2(sum(l.amount for l in lines if l.amount is not None))
        label = rules.agreement_label(ref.agreements, sup)
        if label == "未匹配协议":
            validation.append(ValidationItem("供应商未匹配协议", f"{sup} 不在协议签订记录中"))
        elif label == "是(版本未知)":
            validation.append(ValidationItem("版本异常", f"{sup} 协议版本列为空或脏数据"))
        version = label if label in rules.KNOWN_VERSIONS else ""
        pass_rate = rules.batch_pass_rate(ref.inspections, sup, report_month)
        coef = rules.coefficient(version, pass_rate)
        results.append(SupplierResult(
            supplier=sup, deduction=deduction, pass_rate=pass_rate, agreement=label,
            coefficient=coef, undertaken=rules.round_undertaken(deduction, coef),
            under_200=deduction < 200, skus=lines))       # 规则23：乘系数前判断

    # 规则3：买家备注复核清单（全部计入 + 列出复核）
    review = []
    for row, fbm in [(r, False) for r in fba_rows] + [(r, True) for r in fbm_rows]:
        if row.buyer_comment.strip() and rules.normalize_reason(row.reason_raw, fbm=fbm):
            agg = dlm_agg.get(row.sku)
            review.append(ReviewLine(row.order_id, row.sku, row.reason_raw,
                                     row.buyer_comment, row.return_time,
                                     agg.default_supplier if agg else "未匹配DLM"))

    # 口径差异校验：领星全原因退货合计 vs DLM 退货量合计，差异率>5% 提示
    total_platform = sum(r.qty for r in fba_rows) + sum(r.qty for r in fbm_rows)
    total_dlm = round(sum(a.return_qty for a in dlm_agg.values()), 2)
    base = max(total_platform, total_dlm)
    if base > 0 and abs(total_platform - total_dlm) / base > 0.05:
        validation.append(ValidationItem(
            "口径差异",
            f"领星全原因退货合计 {total_platform} vs DLM 退货量合计 {total_dlm}，"
            f"差异率 {abs(total_platform - total_dlm) / base:.1%}，请检查导出期间是否一致"))

    data = ReportData(report_month)
    data.suppliers = sorted([r for r in results if not r.under_200],
                            key=lambda r: (-r.deduction, r.supplier))
    data.low200 = sorted([r for r in results if r.under_200], key=lambda r: r.supplier)
    data.review = review
    # 剔除占比异常告警：正常情况剔除的都是不再交货的老品（少量）；
    # 若剔除占比过高，大概率是采购入库单导出范围太窄/传错文件 → 只报数量提醒核对
    kept = sum(len(r.skus) for r in results)
    if dropped and kept + dropped and dropped / (kept + dropped) >= 0.3:
        validation.insert(0, ValidationItem(
            "入库单覆盖不足",
            f"有质量退货的 SKU 中 {dropped} 个（占 {dropped / (kept + dropped):.0%}）在采购入库单里"
            f"无交货记录已按老品剔除——若非预期，请检查采购入库单是否导出了 2026-01-01 至报告月末 的完整数据"))
    data.validation = validation
    data.batch_matrix = rules.batch_matrix(ref.inspections)   # 第 4 节「供应商批次合格率」sheet
    return data


def run_month(report_month, fba_path, fbm_path, dlm_path, inbound_path,
              ref: ReferenceData, out_dir: str, store: Store | None = None,
              pdf: bool = False) -> RunSummary:
    data = build_report_data(
        report_month,
        loaders.load_fba(fba_path), loaders.load_fbm(fbm_path), loaders.load_dlm(dlm_path),
        loaders.load_inbound(inbound_path), ref)
    if store is not None:
        store.upsert_month(report_month, data.suppliers + data.low200)
        data.quarterly = store.quarter_rows(report_month)   # 历史+本月合成
    path, low_path = write_report(out_dir, data)
    summary = RunSummary(
        report_month=report_month, file_name=Path(path).name, report_path=str(path),
        supplier_count=len(data.suppliers) + len(data.low200), low200_count=len(data.low200),
        review_count=len(data.review), validation_count=len(data.validation),
        missing_price_count=sum(1 for v in data.validation if v.kind == "缺单价"),
        missing_price_skus=[v.detail for v in data.validation if v.kind == "缺单价"],
        low200_file_name=Path(low_path).name, low200_file_path=str(low_path))
    if pdf:
        try:
            pdf_zip, n = generate_supplier_pdfs(out_dir, data)
            summary.pdf_zip_path, summary.pdf_count = pdf_zip, n
        except Exception as e:      # soffice 缺失等 → 报告照常交付，PDF 标记失败
            summary.pdf_zip_path, summary.pdf_count = "", 0
            data.validation.append(ValidationItem(
                "PDF生成失败", f"供应商 PDF 未生成：{e}"))
            write_report(out_dir, data)   # 把校验项补进报告
    return summary


def generate_supplier_pdfs(out_dir: str, data: ReportData) -> tuple[str, int]:
    """每个供应商结算清单 → 单 sheet xlsx → LibreOffice 转 PDF（年-月+供应商.pdf）→ 打包 zip。"""
    import shutil
    import subprocess
    import tempfile
    import zipfile

    y, m = data.report_month.split("-")
    out = Path(out_dir)
    pdf_dir = out / "pdfs" / data.report_month
    pdf_dir.mkdir(parents=True, exist_ok=True)
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError("未找到 soffice/libreoffice（服务器需安装 libreoffice-calc）")
    # 用户拍板 2026-08-20：PDF 只给 ≥200 的正式结算供应商生成（低于200 不出 PDF）
    suppliers = sorted(data.suppliers, key=lambda s: s.supplier)
    made: list[Path] = []
    with tempfile.TemporaryDirectory() as td:
        for s in suppliers:
            tmp_xlsx = str(Path(td) / "one.xlsx")
            write_supplier_workbook(tmp_xlsx, data, s)
            subprocess.run([soffice, "--headless", "--convert-to", "pdf",
                            "--outdir", td, tmp_xlsx],
                           check=True, capture_output=True, timeout=120)
            src = Path(td) / "one.pdf"
            if not src.exists():
                continue
            dst = pdf_dir / f"{y}-{int(m)}{s.supplier}.pdf"
            shutil.move(str(src), dst)
            made.append(dst)
    zip_path = out / f"{y}年{int(m)}月供应商结算单PDF.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in made:
            zf.write(p, p.name)
    return str(zip_path), len(made)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="供应商质量退货月度报告 CLI")
    ap.add_argument("--month", required=True)
    ap.add_argument("--fba", required=True)
    ap.add_argument("--fbm", required=True)
    ap.add_argument("--dlm", required=True)
    ap.add_argument("--inbound", required=True)
    ap.add_argument("--inspection")
    ap.add_argument("--agreements")
    ap.add_argument("--db", default=str(Path(__file__).resolve().parent.parent / "data" / "app.db"))
    ap.add_argument("--out", default=str(Path(__file__).resolve().parent.parent / "data" / "reports"))
    args = ap.parse_args(argv)

    store = Store(args.db)
    ref = store.load_reference()                # 只更新传入的部分，不清空另一部分
    if args.inspection:
        ref.inspections = loaders.load_inspection(args.inspection)
    if args.agreements:
        ref.agreements = loaders.load_agreements(args.agreements)
    if args.inspection or args.agreements:
        store.save_reference(ref)

    summary = run_month(args.month, args.fba, args.fbm, args.dlm, args.inbound,
                        ref, args.out, store=store)
    print(json.dumps(asdict(summary), ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
