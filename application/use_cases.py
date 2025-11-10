"""Casos de uso da aplicação."""
from typing import List
from domain.entities import Imovel


class ProcessarImoveisUseCase:
    """Caso de uso: processar lista de imóveis e gerar saída."""
    
    def __init__(self, reader, writer):
        self.reader = reader
        self.writer = writer
    
    def execute(self, input_path: str, sheet_name: str, output_path: str) -> int:
        imoveis = self.reader.ler_imoveis(input_path, sheet_name)
        
        if not imoveis:
            return 0
        
        blocos = [imovel.to_text_block() for imovel in imoveis]
        self.writer.escrever_blocos(output_path, blocos)
        
        return len(blocos)
