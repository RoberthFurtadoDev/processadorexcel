"""Serviços de domínio para formatação."""
from decimal import Decimal, InvalidOperation


class FormatadorNumerico:
    """Formata números para padrão brasileiro."""
    
    @staticmethod
    def format_decimal_pt(val_str: str, places: int = 2) -> str:
        if not val_str:
            return ""
        try:
            d = Decimal(str(val_str).replace(",", "."))
            q = d.quantize(Decimal(10) ** -places)
            s = format(q, "f")
            s = s.rstrip("0").rstrip(".")
            s = s.replace(".", ",")
            return s
        except Exception:
            return val_str.replace(".", ",")
    
    @staticmethod
    def iptu_as_text(value) -> str:
        if value is None:
            return ""
        if isinstance(value, str):
            return value.strip()
        if isinstance(value, int):
            return str(value)
        try:
            d = Decimal(str(value))
            s = format(d, "f").rstrip("0").rstrip(".")
            return s if s else "0"
        except (InvalidOperation, ValueError):
            return str(value).strip()


class FormatadorRotulos:
    """Formata pares quantidade + rótulo."""
    
    NORM_MAP = {
        "wc": "wc", "wcs": "wcs", "wc.": "wc",
        "quartos": "quartos", "quarto": "quarto",
        "área de serv.": "área de serv.",
        "area de serv.": "área de serv.",
        "área de serviço": "área de serv.",
        "sala": "sala", "salas": "salas",
        "cozinha": "cozinha", "varanda": "varanda",
        "quartos qts": "quartos"
    }
    
    @classmethod
    def pair_qty_label(cls, qtd: str, label_raw: str, default_label: str) -> str:
        q = FormatadorNumerico.format_decimal_pt(qtd, 0)
        if not q:
            return ""
        
        label = cls.NORM_MAP.get(label_raw.strip().lower(), label_raw.strip())
        
        if q == "1":
            if label.endswith("s") and label not in ("wcs", "áreas de serv."):
                label = label[:-1]
            if label == "wcs":
                label = "wc"
        else:
            if label == "wc":
                label = "wcs"
        
        return f"{q} {label}"
