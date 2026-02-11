"""
Schema models for staged Excel workbook generation.
"""

from typing import Any, Dict, List, Literal, Optional

from pydantic import BaseModel, Field

TemplateName = Literal["executive_dashboard", "finance", "operations"]
ColumnType = Literal["text", "number", "currency", "percentage", "date", "formula"]
ChartType = Literal["bar", "line", "pie"]


class WorkbookColumnSpec(BaseModel):
    name: str
    type: ColumnType
    formula: Optional[str] = None
    sample_values: List[Any] = Field(default_factory=list)


class WorkbookSheetSchema(BaseModel):
    name: str
    columns: List[WorkbookColumnSpec]
    row_count: int = Field(default=12, ge=1, le=500)


class WorkbookDataSchema(BaseModel):
    sheets: List[WorkbookSheetSchema]
    relationships: List[str] = Field(default_factory=list)


class WorkbookLayoutPlan(BaseModel):
    header_row: int = Field(default=4, ge=1, le=50)
    data_start_row: int = Field(default=5, ge=2, le=100)
    freeze_panes: Dict[str, str] = Field(default_factory=dict)
    auto_filter_sheets: List[str] = Field(default_factory=list)
    column_widths: Dict[str, Dict[str, float]] = Field(default_factory=dict)


class WorkbookChartPlan(BaseModel):
    sheet: str
    chart_type: ChartType
    title: str
    data_sheet: str
    category_range: str
    value_range: str
    anchor: str = "F4"


class WorkbookContentPlan(BaseModel):
    workbook_title: str
    workbook_subtitle: str = ""
    summary_points: List[str] = Field(default_factory=list)
    chart_plans: List[WorkbookChartPlan] = Field(default_factory=list)


class WorkbookStylePlan(BaseModel):
    theme_name: str
    primary_color: str
    secondary_color: str
    accent_color: str
    font_name: str = "Calibri"
    header_font_size: int = 12
    body_font_size: int = 11
    zebra_fill_color: str = "F7F9FC"


class WorkbookQcPlan(BaseModel):
    block_on_critical: bool = True
    minimum_visual_score: float = Field(default=0.78, ge=0.0, le=1.0)
    max_refinement_iterations: int = Field(default=2, ge=0, le=5)


class WorkbookSpec(BaseModel):
    intent: str
    template: TemplateName
    data_schema: WorkbookDataSchema
    layout_plan: WorkbookLayoutPlan
    content_plan: WorkbookContentPlan
    style_plan: WorkbookStylePlan
    qc_plan: WorkbookQcPlan


class WorkbookSpecSummary(BaseModel):
    template: TemplateName
    sheet_count: int
    column_count: int
    has_charts: bool
    refinement_iterations: int


class GenerationScore(BaseModel):
    visual_score: float = Field(ge=0.0, le=1.0)
    structure_score: float = Field(ge=0.0, le=1.0)
    formula_score: float = Field(ge=0.0, le=1.0)
    overall_score: float = Field(ge=0.0, le=1.0)


class ExcelGenerateRequest(BaseModel):
    prompt: str = Field(min_length=1, max_length=8000)
    template: TemplateName = "executive_dashboard"
    filename: Optional[str] = None
    include_sample_data: bool = True
    include_charts: bool = True
    max_rows_per_sheet: int = Field(default=20, ge=3, le=200)
    minimum_visual_score: float = Field(default=0.78, ge=0.0, le=1.0)
    max_refinement_iterations: int = Field(default=2, ge=0, le=5)
    block_on_critical_qc: bool = True


class ExcelGenerateResponse(BaseModel):
    status: Literal["ok", "blocked", "error"]
    message: str
    fileId: Optional[str] = None
    downloadUrl: Optional[str] = None
    artifact: Optional[Dict[str, Any]] = None
    workbookSpec: Optional[WorkbookSpecSummary] = None
    generationScore: Optional[GenerationScore] = None
    qcReport: Optional[Dict[str, Any]] = None
    repairsApplied: int = 0

