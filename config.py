"""Configurações da aplicação."""
import os

CONFIG = {
    "default_input": os.getenv("INPUT_FILE", "Teste tit2.xlsx"),
    "default_sheet": os.getenv("INPUT_SHEET", "Planilha1"),
    "output_file": os.getenv("OUTPUT_FILE", "dados_formatados.xlsx"),
    "col_letter": "A",
    "col_width_chars": 28,
    "base_row_height": 18,
    "extra_padding_lines": 2,
    "row_gap": 0,
}
