"""合成 xlsx/csv 工厂：列名与真实导出文件一致（见 plan.md 列名速查）。
工厂第一个参数是目标路径（str|Path），返回同一路径 str。"""
from __future__ import annotations

import csv
from pathlib import Path

import openpyxl

FBA_HEADER = ["店铺", "国家", "品名", "sku", "分类", "品牌", "订单号", "商品名称", "MSKU",
              "ASIN", "FNSKU", "退货数量", "发货仓库编号", "库存属性", "退货原因", "状态",
              "LPN编号", "买家备注", "退货时间", "订购时间", "标签", "备注"]
FBM_HEADER = ["订单号", "Prime订单", "A-to-Z索赔", "商品名称", "ASIN", "MSKU", "原始数量", "品名",
              "SKU", "分类", "品牌", "退货数量", "退货原因", "退货状态", "退货时间", "标签", "备注"]
DLM_HEADER1 = ["SKU", "SKU名称", "默认供应商", "其他供应商", "销量", "退货量",
               "库存属性情况分析", "", "退货原因类别归纳", ""]
DLM_HEADER2 = ["", "", "", "", "", "", "可销售", "可销售占比", "客户问题", "客户问题占比"]
INBOUND_HEADER = ["入库日期", "创建日期", "单据编号", "供应商", "物料编码", "应收数量",
                  "实收数量", "总金额", "含税单价", "单价"]
INSPECTION_HEADER = ["单据编号", "单据时间", "供应商", "验货地址", "月份", "SKU名称", "质检结果"]
AGREEMENT_HEADER = ["序号", "供应商代码", "供应商名称", "sourcing", "合作状态", "类型",
                    "框架合同", "框架合同补充协议", "质量协议", "", "廉洁协议", "是否函调", "备注"]


def _save(path, rows, sheet_name="Sheet1"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    for r in rows:
        ws.append(list(r))
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return str(path)


def _patch(header, cols, values):
    row = [""] * len(header)
    for col, v in zip(cols, values):
        row[header.index(col)] = v
    return row


def make_fba_file(path, data_rows):
    """data_rows: (订单号, sku, 退货数量, 退货原因, 退货时间, 买家备注)"""
    cols = ("订单号", "sku", "退货数量", "退货原因", "退货时间", "买家备注")
    return _save(path, [FBA_HEADER] + [_patch(FBA_HEADER, cols, r) for r in data_rows],
                 sheet_name="sheet1")


def make_fbm_file(path, data_rows):
    """data_rows: (订单号, SKU, 退货数量, 退货原因, 退货时间, 备注)"""
    cols = ("订单号", "SKU", "退货数量", "退货原因", "退货时间", "备注")
    return _save(path, [FBM_HEADER] + [_patch(FBM_HEADER, cols, r) for r in data_rows],
                 sheet_name="sheet1")


def make_dlm_file(path, data_rows):
    """data_rows: (SKU, 默认供应商, 其他供应商, 销量, 退货量)。两行表头。"""
    cols = ("SKU", "默认供应商", "其他供应商", "销量", "退货量")
    return _save(path, [DLM_HEADER1, DLM_HEADER2]
                 + [_patch(DLM_HEADER1, cols, r) for r in data_rows])


def make_inbound_file(path, data_rows):
    """data_rows: (入库日期, 供应商, 物料编码, 实收数量, 单价)"""
    cols = ("入库日期", "供应商", "物料编码", "实收数量", "单价")
    return _save(path, [INBOUND_HEADER] + [_patch(INBOUND_HEADER, cols, r) for r in data_rows])


def make_inspection_file(path, data_rows):
    """data_rows: (供应商, 月份, 质检结果)"""
    cols = ("供应商", "月份", "质检结果")
    return _save(path, [INSPECTION_HEADER] + [_patch(INSPECTION_HEADER, cols, r) for r in data_rows],
                 sheet_name="26年验货原始数据")


def make_agreements_csv(path, data_rows):
    """data_rows: (供应商名称, 质量协议, 版本)。模拟飞书导出的 [row=N] 前缀。"""
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(AGREEMENT_HEADER)
        for i, (name, signed, version) in enumerate(data_rows, start=1):
            w.writerow([i, f"VEN{i:05d}", name, "某人", "合格供应商", "OEM",
                        "是", "", signed, version, "是", "是", ""])
    text = Path(path).read_text(encoding="utf-8")
    with open(path, "w", encoding="utf-8", newline="") as f:
        for i, ln in enumerate(text.splitlines()):
            f.write(f"[row={3 + i}] {ln}\n")
    return str(path)


def make_agreements_xlsx(path, data_rows):
    """data_rows: (供应商名称, 质量协议, 质量协议版本)"""
    return _save(path, [["供应商名称", "质量协议", "质量协议版本"]] + [list(r) for r in data_rows])


def make_inbound_file_with_tax(path, data_rows):
    """data_rows: (入库日期, 供应商, 物料编码, 实收数量, 单价, 含税单价)"""
    cols = ("入库日期", "供应商", "物料编码", "实收数量", "单价", "含税单价")
    return _save(path, [INBOUND_HEADER] + [_patch(INBOUND_HEADER, cols, r) for r in data_rows])
