"""Entidades de domínio."""
from dataclasses import dataclass
from typing import Optional


@dataclass
class Imovel:
    """Representa um imóvel com seus atributos."""
    empreendimento: str = ""
    tipo: str = ""
    area_total: str = ""
    area_privativa: str = ""
    area_terreno: str = ""
    quartos: str = ""
    vaga: str = ""
    varanda: str = ""
    area_servico: str = ""
    banheiros: str = ""
    salas: str = ""
    cozinhas: str = ""
    descricao: str = ""
    iptu: str = ""
    matricula: str = ""
    oficio: str = ""

    def to_text_block(self) -> str:
        """Converte o imóvel para bloco de texto formatado."""
        parts = []
        
        if self.empreendimento:
            parts.append(f"{self.empreendimento}.")
        if self.tipo:
            parts.append(f"{self.tipo},")
        if self.area_total:
            parts.append(f"{self.area_total} m² de área total,")
        if self.area_privativa:
            parts.append(f"{self.area_privativa} m² de área privativa,")
        if self.area_terreno:
            parts.append(f"{self.area_terreno} m² de área do terreno,")
        if self.vaga:
            parts.append(f"{self.vaga},")
        if self.varanda:
            parts.append(f"{self.varanda},")
        if self.quartos:
            parts.append(f"{self.quartos},")
        if self.area_servico:
            parts.append(f"{self.area_servico},")
        if self.banheiros:
            parts.append(f"{self.banheiros},")
        if self.salas:
            parts.append(f"{self.salas},")
        if self.cozinhas:
            parts.append(f"{self.cozinhas},")
        
        first_line = " ".join(parts).replace(" ,", ",").replace(",,", ",").strip()
        if first_line.endswith(","):
            first_line = first_line[:-1]
        if not first_line.endswith("."):
            first_line += "."
        first_line = first_line.replace("..", ".")
        
        desc_block = f"\n\n{self.descricao}" if self.descricao else ""
        
        return f"{first_line}{desc_block}\n\nIPTU: {self.iptu}\n\nMatrícula: {self.matricula} Ofício: {self.oficio}."
