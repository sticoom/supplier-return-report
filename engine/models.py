"""共享数据类：模块间契约载体。无逻辑、无 IO。"""
from __future__ import annotations

from dataclasses import dataclass, field

QUALITY_REASON_CODES = ("DEFECTIVE", "MISSING_PARTS", "QUALITY_UNACCEPTABLE")


@dataclass(frozen=True)
class ReturnOrderRow:
    order_id: str
    sku: str
    qty: int
    reason_raw: str
    return_time: str
    buyer_comment: str


@dataclass(frozen=True)
class DlmRow:
    sku: str
    default_supplier: str
    other_supplier: str
    sales_qty: float
    return_qty: float


@dataclass
class DlmAgg:
    sku: str
    default_supplier: str
    other_supplier: str
    sales_qty: float
    return_qty: float


@dataclass(frozen=True)
class InboundRow:
    date: str          # 'YYYY-MM-DD'
    supplier: str
    sku: str
    qty: float
    unit_price: float  # 不含税单价


@dataclass(frozen=True)
class InspectionBatch:
    supplier: str      # strip 后全名
    month: str         # 'YYYY-MM'
    result: str        # '合格' / '不合格'
    count: int = 1     # 批数（汇总页导入时>1，原始明细恒为1）


@dataclass(frozen=True)
class AgreementInfo:
    supplier: str
    signed: bool
    version: str       # '' / 'V2版本' / 'V3版' / 'V4版'（脏值归 ''）


@dataclass
class SkuLine:
    sku: str
    sales_qty: float | None
    return_qty: float | None
    qty_defective: float
    qty_missing_parts: float
    qty_quality_unacceptable: float
    rate: float | None
    unit_price: float | None
    amount: float | None
    note: str


@dataclass
class SupplierResult:
    supplier: str
    deduction: float
    pass_rate: float | None
    agreement: str
    coefficient: float
    undertaken: int
    under_200: bool
    skus: list[SkuLine] = field(default_factory=list)


@dataclass
class ReviewLine:
    order_id: str
    sku: str
    reason_raw: str
    buyer_comment: str
    return_time: str
    supplier: str


@dataclass
class ValidationItem:
    kind: str
    detail: str


@dataclass
class QuarterRow:
    month: str
    supplier: str
    deduction: float
    undertaken: int
    is_subtotal: bool


@dataclass
class ReferenceData:
    inspections: list[InspectionBatch] = field(default_factory=list)
    agreements: list[AgreementInfo] = field(default_factory=list)


@dataclass
class ReportData:
    report_month: str
    suppliers: list[SupplierResult] = field(default_factory=list)
    low200: list[SupplierResult] = field(default_factory=list)
    quarterly: list[QuarterRow] = field(default_factory=list)
    review: list[ReviewLine] = field(default_factory=list)
    validation: list[ValidationItem] = field(default_factory=list)
    # 第 4 节「供应商批次合格率」sheet 数据源：供应商全名 → 'YYYY-MM' → (总批数, 不合格批数)
    batch_matrix: dict[str, dict[str, tuple[int, int]]] = field(default_factory=dict)


@dataclass
class RunSummary:
    report_month: str
    file_name: str
    report_path: str
    supplier_count: int
    low200_count: int
    review_count: int
    validation_count: int
    missing_price_count: int
