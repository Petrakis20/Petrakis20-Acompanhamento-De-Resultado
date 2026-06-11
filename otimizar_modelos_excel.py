import re
import zipfile
from pathlib import Path


MODELOS = [
    {
        "origem": Path("NOVO_FORMATO/Atualização de Acompanhamento LP versão 2.xlsx"),
        "destino": Path("NOVO_FORMATO/otimizados/Atualização de Acompanhamento LP versão 2 - otimizado.xlsx"),
        "sheets": {
            "xl/worksheets/sheet1.xml": {"max_row": 206, "dimension": "A1:Q206"},
        },
    },
    {
        "origem": Path("NOVO_FORMATO/Atualização de Acompanhamento LR versão 2.xlsx"),
        "destino": Path("NOVO_FORMATO/otimizados/Atualização de Acompanhamento LR versão 2 - otimizado.xlsx"),
        "sheets": {
            "xl/worksheets/sheet2.xml": {"max_row": 197, "dimension": "A1:R197"},
        },
    },
]

ROW_RE = re.compile(r'<row\b[^>]*\br=\"(\d+)\"[^>]*?(?:/>|>(?:.*?</row>))', re.DOTALL)
DIMENSION_RE = re.compile(r'<dimension ref="[^"]+"')


def compactar_sheet(xml, max_row, dimension):
    def substituir_linha(match):
        row_number = int(match.group(1))
        return match.group(0) if row_number <= max_row else ""

    xml = ROW_RE.sub(substituir_linha, xml)
    return DIMENSION_RE.sub(f'<dimension ref="{dimension}"', xml, count=1)


def otimizar_modelo(origem, destino, sheets):
    destino.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(origem, "r") as zin, zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in sheets:
                regra = sheets[item.filename]
                xml = data.decode("utf-8")
                data = compactar_sheet(xml, regra["max_row"], regra["dimension"]).encode("utf-8")
            zout.writestr(item, data)


def main():
    for modelo in MODELOS:
        otimizar_modelo(modelo["origem"], modelo["destino"], modelo["sheets"])
        print(f"Gerado: {modelo['destino']}")


if __name__ == "__main__":
    main()
