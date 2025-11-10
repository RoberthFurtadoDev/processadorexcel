"""Adaptador de escrita Excel."""
import math
from typing import List
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side


class ExcelWriter:
    """Escreve blocos de texto em planilhas Excel."""
    
    def __init__(self, config: dict):
        self.col_letter = config.get("col_letter", "A")
        self.col_width = config.get("col_width_chars", 28)
        self.base_row_height = config.get("base_row_height", 18)
        self.extra_padding = config.get("extra_padding_lines", 2)
        self.row_gap = config.get("row_gap", 0)
    
    def escrever_blocos(self, output_path: str, blocos: List[str]):
        wb = Workbook()
        ws = wb.active
        ws.title = "Registros"
        ws.column_dimensions[self.col_letter].width = self.col_width
        
        thin = Side(style="thin", color="000000")
        border_all = Border(top=thin, left=thin, right=thin, bottom=thin)
        
        row_ptr = 1
        for bloco in blocos:
            cell = ws[f"{self.col_letter}{row_ptr}"]
            cell.value = bloco
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border_all
            ws.row_dimensions[row_ptr].height = self._estimate_row_height(bloco)
            row_ptr += 1 + self.row_gap
        
        wb.save(output_path)
    
    def _estimate_row_height(self, text: str) -> float:
        total_lines = 0
        for para in text.split("\n"):
            total_lines += max(1, math.ceil(len(para) / max(1, self.col_width))) if para else 1
        total_lines += self.extra_padding
        return total_lines * self.base_row_height
