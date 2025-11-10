"""Ponto de entrada da aplicação."""
import sys
import os
from config import CONFIG
from application.use_cases import ProcessarImoveisUseCase
from infrastructure.excel_reader import ExcelReader
from infrastructure.excel_writer import ExcelWriter


def main():
    input_file = sys.argv[1] if len(sys.argv) >= 2 else CONFIG["default_input"]
    sheet_name = sys.argv[2] if len(sys.argv) >= 3 else CONFIG["default_sheet"]
    output_file = CONFIG["output_file"]
    
    if not os.path.exists(input_file):
        print(f"Erro: O arquivo de entrada '{input_file}' não foi encontrado.")
        return
    
    reader = ExcelReader()
    writer = ExcelWriter(CONFIG)
    use_case = ProcessarImoveisUseCase(reader, writer)
    
    try:
        total = use_case.execute(input_file, sheet_name, output_file)
        if total > 0:
            print(f"Arquivo '{output_file}' gerado com {total} cards em A1, A2, A3, ...")
        else:
            print("Nenhum registro foi formatado.")
    except Exception as e:
        print(f"Erro ao processar: {e}")
        raise


if __name__ == "__main__":
    main()
