"""openpyxl 报告输出：复刻手工模板版式。"""
from __future__ import annotations

import re
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side

from engine.models import ReportData, SupplierResult

BOLD = Font(bold=True)
TITLE_FONT = Font(bold=True, size=14)
THIN = Border(left=Side(style="thin"), right=Side(style="thin"),
              top=Side(style="thin"), bottom=Side(style="thin"))
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

BAD_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")   # Excel sheet 名非法字符
FIXED_SHEET_NAMES = ("供应商费用明细", "低于200清单", "季度累计", "买家备注复核清单",
                     "供应商批次合格率", "数据校验")
SUMMARY_HEADER = ["序号", "供应商", "应扣金额", "批次合格率", "是否签署质量协议(版本)",
                  "承担比例(考核系数)", "应承担金额", "备注"]
REASON_SUBS = ["DEFECTIVE \n(存在瑕疵)", "MISSING_PARTS\n (部分零件缺失)",
               "QUALITY_UNACCEPTABLE\n (质量未达到期望)"]
NOTE_TEXT = ("备注：数据来源于领星系统，亚马逊后台数据；如对数据有疑问，请及时联系德拉姆品质部。"
             "考核依据：依据《质量保证协议》相应版本条款。")
SIGN_ROWS = [("采购部", "审核及优化意见简述："), ("品质部", "审核及优化意见简述："),
             ("供应链中心", "审核及优化意见简述："), ("总经理", "最终意见：")]


def short_name(supplier: str) -> str:
    s = supplier.strip()
    for suffix in ("有限公司", "有限责任公司", "公司"):
        if s.endswith(suffix) and len(s) > len(suffix):
            s = s[: -len(suffix)]
            break
    head = s[:5]                      # 行政区划前缀（东莞市/台州/石家庄市…）在头部
    for marker in ("市", "州"):        # 先找「市」（如苏州市），再退回「州」（如台州）
        idx = head.rfind(marker)
        if 1 <= idx < len(s) - 1:
            return s[idx + 1:]
    return s or supplier


def _sheet_title(base: str, used: set[str]) -> str:
    name = BAD_SHEET_CHARS.sub("-", base)[:31] or "供应商"
    candidate, n = name, 2
    while candidate in used:
        suffix = f"({n})"
        candidate = name[: 31 - len(suffix)] + suffix
        n += 1
    used.add(candidate)
    return candidate


def _style_header(ws, row, headers, start_col=1):
    for j, h in enumerate(headers, start=start_col):
        c = ws.cell(row=row, column=j, value=h)
        c.font = BOLD
        c.alignment = CENTER
        c.border = THIN


def _quarter_label(month: str) -> str:
    y, m = month.split("-")           # 刻意内联：保持 report 不依赖 rules（任务解耦）
    return f"{y}-Q{(int(m) - 1) // 3 + 1}"


def _summary_sheet(wb: Workbook, data: ReportData, only_low200: bool) -> None:
    ws = wb.create_sheet("低于200清单" if only_low200 else "供应商费用明细")
    header = (SUMMARY_HEADER + ["所属季度"]) if only_low200 else SUMMARY_HEADER
    _style_header(ws, 1, header)
    pool = data.low200 if only_low200 else (data.suppliers + data.low200)
    ordered = sorted(pool, key=lambda s: (-s.deduction, s.supplier))
    for i, s in enumerate(ordered, start=2):
        note = "本月无验货" if s.pass_rate is None else ""
        vals = [i - 1, s.supplier, s.deduction,
                s.pass_rate, s.agreement, s.coefficient, s.undertaken, note]
        if only_low200:
            vals.append(_quarter_label(data.report_month))
        for j, v in enumerate(vals, start=1):
            c = ws.cell(row=i, column=j, value=v)
            c.border = THIN
        ws.cell(row=i, column=3).number_format = "0.00"
        if s.pass_rate is not None:
            ws.cell(row=i, column=4).number_format = "0.0000%"
        ws.cell(row=i, column=6).number_format = "0.##"
    ws.column_dimensions["B"].width = 34


def _supplier_sheet(wb: Workbook, used: set[str], data: ReportData,
                    s: SupplierResult) -> None:
    ws = wb.create_sheet(_sheet_title(short_name(s.supplier), used))
    y, m = data.report_month.split("-")
    ws.merge_cells("A1:K1")
    ws["A1"] = f" {y} 年 {int(m)} 月供应商质量退货金额汇总表"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = f"供应商名称：{s.supplier}"
    for col, w in {"A": 6, "B": 20, "C": 8, "D": 8, "E": 13, "F": 15, "G": 17,
                   "H": 11, "I": 13, "J": 14, "K": 18}.items():
        ws.column_dimensions[col].width = w

    main = ["序号", "SKU名称", "销量", "退货量", "质量退货量（按亚马逊平台退货描述）",
            None, None, "质量退货率", "产品采购单价\n（元）", "质量退货金额\n（元）", "备注"]
    for j, h in enumerate(main, start=1):
        if h is not None:
            _style_header(ws, 3, [h], start_col=j)
    for j, h in enumerate(REASON_SUBS, start=5):
        _style_header(ws, 4, [h], start_col=j)
    for rng in ("A3:A4", "B3:B4", "C3:C4", "D3:D4", "E3:G3",
                "H3:H4", "I3:I4", "J3:J4", "K3:K4"):
        ws.merge_cells(rng)

    r = 5
    for i, ln in enumerate(s.skus, start=1):        # skus 已只含质量退货量>0 的行
        vals = [i, ln.sku, ln.sales_qty, ln.return_qty, ln.qty_defective,
                ln.qty_missing_parts, ln.qty_quality_unacceptable,
                ln.rate, ln.unit_price, ln.amount, ln.note or None]
        for j, v in enumerate(vals, start=1):
            c = ws.cell(row=r, column=j, value=v)
            c.border = THIN
        ws.cell(row=r, column=8).number_format = "0.00%"
        ws.cell(row=r, column=9).number_format = "0.00##"
        ws.cell(row=r, column=10).number_format = "0.00"
        r += 1
    last = r - 1 if r > 5 else 5

    r_sum = r
    ws.cell(row=r_sum, column=1, value="统计金额：").font = BOLD
    ws.cell(row=r_sum, column=10, value=f"=SUM(J5:J{last})").number_format = "0.00"
    r_coef = r_sum + 1
    ws.cell(row=r_coef, column=1, value="是否签署最新质量协议").font = BOLD
    ws.cell(row=r_coef, column=5, value=s.agreement)
    ws.cell(row=r_coef, column=8, value="考核系数：").font = BOLD
    ws.cell(row=r_coef, column=10, value=s.coefficient).number_format = "0.##"
    r_pass = r_coef + 1
    ws.cell(row=r_pass, column=1, value="当月检验合格率：").font = BOLD
    c = ws.cell(row=r_pass, column=5, value="/" if s.pass_rate is None else s.pass_rate)
    if s.pass_rate is not None:
        c.number_format = "0.0000%"
    ws.cell(row=r_pass, column=8, value="考核金额：").font = BOLD
    ws.cell(row=r_pass, column=10, value=f"=J{r_coef}*J{r_sum}").number_format = "0"
    ws.cell(row=r_pass + 1, column=1, value=NOTE_TEXT)
    for k, (dept, label) in enumerate(SIGN_ROWS):
        rr = r_pass + 3 + k
        ws.cell(row=rr, column=1, value=dept).font = BOLD
        ws.cell(row=rr, column=2, value=label)
        ws.cell(row=rr, column=7, value="签字：")
        ws.cell(row=rr, column=9, value="日期：")


def _quarterly_sheet(wb: Workbook, data: ReportData) -> None:
    q = wb.create_sheet("季度累计")
    _style_header(q, 1, ["月份", "供应商", "应扣金额", "应承担金额", "备注"])
    for i, row in enumerate(data.quarterly, start=2):
        vals = [row.month, row.supplier, row.deduction, row.undertaken,
                "季度小计" if row.is_subtotal else ""]
        for j, v in enumerate(vals, start=1):
            c = q.cell(row=i, column=j, value=v)
            if row.is_subtotal:
                c.font = BOLD
        q.cell(row=i, column=3).number_format = "0.00"


def _review_sheet(wb: Workbook, data: ReportData) -> None:
    rv = wb.create_sheet("买家备注复核清单")
    _style_header(rv, 1, ["订单号", "SKU", "退货原因", "买家备注", "退货时间", "供应商"])
    for i, row in enumerate(data.review, start=2):
        for j, v in enumerate([row.order_id, row.sku, row.reason_raw,
                               row.buyer_comment, row.return_time, row.supplier], start=1):
            rv.cell(row=i, column=j, value=v)


def _batch_matrix_sheet(wb: Workbook, data: ReportData) -> None:
    """第 4 节「供应商批次合格率」：供应商×月份矩阵（总批数/不合格批数/批次合格率），
    供应商用全名，品质部可从本表复制粘贴更新 wiki《2026年验货数据报表》的「供应商」sheet。"""
    bm = wb.create_sheet("供应商批次合格率")
    months = sorted({mo for cells in data.batch_matrix.values() for mo in cells})
    _style_header(bm, 1, ["供应商"])
    for j, mo in enumerate(months):
        c0 = 2 + j * 3                          # 每月 3 子列：总批数/不合格批数/批次合格率
        _style_header(bm, 1, [mo], start_col=c0)
        _style_header(bm, 2, ["总批数", "不合格批数", "批次合格率"], start_col=c0)
        bm.merge_cells(start_row=1, start_column=c0, end_row=1, end_column=c0 + 2)
    bm.merge_cells("A1:A2")
    for i, sup in enumerate(sorted(data.batch_matrix), start=3):
        c = bm.cell(row=i, column=1, value=sup)
        c.border = THIN
        for j, mo in enumerate(months):
            cell = data.batch_matrix[sup].get(mo)
            if cell is None:                     # 该月无验货 → 留空
                continue
            total, failed = cell
            for k, v in enumerate((total, failed, (total - failed) / total)):
                cc = bm.cell(row=i, column=2 + j * 3 + k, value=v)
                cc.border = THIN
            bm.cell(row=i, column=2 + j * 3 + 2).number_format = "0.0000%"
    if months:
        bm.cell(row=len(data.batch_matrix) + 4, column=1,
                value="备注：从验货原始数据计算，批次合格率＝合格批数÷总批数（供应商全名）。")
    bm.column_dimensions["A"].width = 34


def _validation_sheet(wb: Workbook, data: ReportData) -> None:
    vd = wb.create_sheet("数据校验")
    _style_header(vd, 1, ["类型", "明细"])
    for i, item in enumerate(data.validation, start=2):
        vd.cell(row=i, column=1, value=item.kind)
        vd.cell(row=i, column=2, value=item.detail)


def write_report(out_dir: str, data: ReportData) -> str:
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    y, m = data.report_month.split("-")
    path = out / f"{y}年{int(m)}月供应商质量退货金额汇总表.xlsx"

    wb = Workbook()
    wb.remove(wb.active)
    used = set(FIXED_SHEET_NAMES)
    _summary_sheet(wb, data, only_low200=False)
    for s in sorted(data.suppliers + data.low200, key=lambda x: x.supplier):
        _supplier_sheet(wb, used, data, s)
    _summary_sheet(wb, data, only_low200=True)
    _quarterly_sheet(wb, data)
    _review_sheet(wb, data)
    _batch_matrix_sheet(wb, data)
    _validation_sheet(wb, data)

    wb.save(path)
    return str(path)
