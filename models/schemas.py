"""
Intermediate data model.

Every extractor must produce data that validates against these schemas.
Every renderer only ever reads from these schemas - never from raw
openpyxl cells. This is the contract that decouples "Excel changed" from
"Word doc generation broke".
"""
from __future__ import annotations
from typing import Optional
from pydantic import BaseModel, Field


class Attribute(BaseModel):
    """One row in a Product/Product-Component's attribute table
    (source: sheet '2. Product Hierarchy', rows tagged 'Attribute' in
    column C, following a 'Product'/'Product Component' anchor row)."""
    entity_name: str
    mandatory: Optional[str] = None
    comments: Optional[str] = None
    values: list[str] = Field(default_factory=list)  # rows tagged "Value" following this attribute
    source_sheet: str
    source_row: int


class ProductComponent(BaseModel):
    """One Product or Product Component node from sheet '2. Product Hierarchy'.
    kind is either 'Product' or 'Product Component' (literal value of column C)."""
    kind: str
    name: str
    attributes: list[Attribute] = Field(default_factory=list)
    source_sheet: str
    source_row: int


class OrderScenarioRow(BaseModel):
    """One row of the Section 3 order-scenario table.

    Source: sheet 5.1 MDM_Order Scenarios.
    Business Scenario is mapped from Service Type; CRM fields come from
    Order Sub Type and Condition; OSS Service Order comes from Order Type.
    """
    business_scenario: str
    crm_order_sub_type: str
    crm_order_condition: str
    oss_service_order: str


class Task(BaseModel):
    """One row inside an order-type task list (source: sheet 6.3 pattern)."""
    seq: Optional[str] = None
    task_name: str
    duration: Optional[float] = None
    cancellation_wo: Optional[str] = None
    it_system: Optional[str] = None
    work_group: Optional[str] = None
    task_condition: Optional[str] = None
    mandatory_attributes: Optional[str] = None
    special_configurations: Optional[str] = None
    description: Optional[str] = None  # joined in from 6.4 when available
    # Seq value inherited from a vertically merged Seq cell (e.g. A6:A7). The
    # task table keeps showing the raw `seq`; the flowchart uses this to put
    # parallel tasks in the same sequence step instead of at the end.
    flow_seq: Optional[str] = None

    # Traceability - always know where a fact came from
    source_sheet: str
    source_row: int


class OrderTypeSection(BaseModel):
    """One stacked table in sheet 6.3 - all tasks for one order type."""
    order_type: str
    product: str  # no default - must be supplied explicitly per pipeline run
    title: str
    source_sheet: str
    source_rows: str
    tasks: list[Task] = Field(default_factory=list)

    @property
    def task_count(self) -> int:
        return len(self.tasks)


class BusinessRuleBlock(BaseModel):
    """One requirement paragraph plus any source tables belonging to it."""
    text: str
    tables: list[ExcelTableData] = Field(default_factory=list)
    source_row: int | None = None


class ExcelTableData(BaseModel):
    """A rectangular Excel table captured for direct rendering into Word."""
    headers: list[str] = Field(default_factory=list)
    rows: list[list[str]] = Field(default_factory=list)
    # Optional top header row for two-level headers: [[label, column_span], ...]
    # covering every column. An entry ["", 1] means the column header is
    # merged vertically across both header rows.
    group_headers: list[list] = Field(default_factory=list)
    source_sheet: str
    source_range: str = ""


class ExtractionResult(BaseModel):
    """Top-level container - one of these per pipeline run."""
    workbook_name: str
    product_name: str
    introduction_text: list[str] = Field(default_factory=list)            # sheet "1. Product Description"
    ai_introduction: list[str] = Field(default_factory=list)                # optional Gemini-enhanced Introduction
    business_rules: list[str] = Field(default_factory=list)               # flattened Requirement Description values
    business_rules_table: ExcelTableData | None = None                    # legacy field; no longer renders the whole requirement sheet
    business_rule_blocks: list[BusinessRuleBlock] = Field(default_factory=list) # requirement text + tables belonging to that requirement
    hierarchy: list[ProductComponent] = Field(default_factory=list)       # sheet "2. Product Hierarchy"
    hierarchy_images: list[str] = Field(default_factory=list)             # embedded images in sheet "2. Product Hierarchy"
    order_scenarios: list[OrderScenarioRow] = Field(default_factory=list) # sheet 5.1 MDM_Order Scenarios
    order_types: list[OrderTypeSection] = Field(default_factory=list)     # sheet 6.3
    warnings: list[str] = Field(default_factory=list)
