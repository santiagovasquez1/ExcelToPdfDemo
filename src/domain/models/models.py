from typing import Any, List
from pydantic import BaseModel


class MergedRangesModel(BaseModel):
    min_row: int
    max_row: int
    min_col: int
    max_col: int

class CellData(BaseModel):
    value: Any
    row: int
    col: int

class MergedRange(BaseModel):
    min_row: int
    max_row: int
    min_col: int
    max_col: int

class SheetData(BaseModel):
    sheet_name: str
    table_data: List[List[CellData]]
    merged_ranges: List[MergedRange]