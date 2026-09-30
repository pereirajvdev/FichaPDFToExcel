import re
import sys
from pathlib import Path

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


COLUNAS_VALORES = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro",
    "13º salário", "Rescisão", "Total",
]

MESES = COLUNAS_VALORES[:12]

PADRAO_DINHEIRO = r"-?\d{1,3}(?:\.\d{3})*,\d{2}"


def extrair_texto_pdf(caminho_pdf: Path) -> str:
    """Extrai o texto de todas as páginas do PDF."""
    paginas = []

    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text() or ""
            paginas.append(texto)

    return "\n".join(paginas)


def normalizar_valor(valor: str) -> float:
    """Converte valor brasileiro, como 1.234,56, para float."""
    return float(valor.replace(".", "").replace(",", "."))


def extrair_dados_servidor(texto: str) -> dict:
    """Extrai os dados do servidor do cabeçalho da ficha."""
    dados = {}

    # Ex.: 000594 MARGARIDA DE JESUS ANDRADE Secretaria: ...
    padrao_servidor = re.search(
        r"(?m)^(\d{6})\s+(.+?)\s+Secretaria:", texto
    )
    if padrao_servidor:
        dados["Matrícula"] = padrao_servidor.group(1)
        dados["Nome"] = padrao_servidor.group(2).strip()

    padrao = re.search(r"Ficha Financeira Resumo Geral do Ano de (\d{4})", texto)
    if padrao:
        dados["Ano"] = int(padrao.group(1))

    padrao = re.search(r"Secretaria:\s*(.*?)\s+Conta:", texto)
    if padrao:
        dados["Secretaria"] = padrao.group(1).strip()

    padrao = re.search(r"Conta:\s*(\S+)\s+Ano:", texto)
    if padrao:
        dados["Conta"] = padrao.group(1)

    padrao = re.search(r"Cargo:\s*(.*?)\s+CTPS:", texto)
    if padrao:
        dados["Cargo"] = padrao.group(1).strip()

    padrao = re.search(r"CTPS:\s*(\S+)\s+Pis/Pasep:", texto)
    if padrao:
        dados["CTPS"] = padrao.group(1)

    padrao = re.search(r"Pis/Pasep:\s*(\S+)\s+CPF:", texto)
    if padrao:
        dados["PIS/PASEP"] = padrao.group(1)

    padrao = re.search(r"CPF:\s*(\S+)\s+Dt\. Nasc\.:", texto)
    if padrao:
        dados["CPF"] = padrao.group(1)

    padrao = re.search(r"Dt\. Nasc\.\s*:\s*(\S+)\s+Dt\. Admissão:\s*(\S+)", texto)
    if padrao:
        dados["Data de nascimento"] = padrao.group(1)
        dados["Data de admissão"] = padrao.group(2)

    return dados


def extrair_rubricas(texto: str) -> list[dict]:
    """Extrai as rubricas que possuem código de 5 dígitos e 15 valores."""
    rubricas = []
    padrao_valor = re.compile(PADRAO_DINHEIRO)

    for linha in texto.splitlines():
        linha = linha.strip()

        if not re.match(r"^\d{5}\s+", linha):
            continue

        codigo = linha[:5]
        restante = linha[5:].strip()
        valores_encontrados = list(padrao_valor.finditer(restante))

        # Cada rubrica da ficha possui exatamente 15 valores.
        if len(valores_encontrados) != 15:
            print(
                f"Aviso: rubrica ignorada por quantidade inesperada de valores: {linha}"
            )
            continue

        descricao = restante[:valores_encontrados[0].start()].strip()
        valores = [normalizar_valor(m.group()) for m in valores_encontrados]

        rubricas.append({
            "Código": codigo,
            "Descrição": descricao,
            **dict(zip(COLUNAS_VALORES, valores)),
        })

    return rubricas


def criar_excel(caminho_saida: Path, dados_servidor: dict, rubricas: list[dict]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Ficha Financeira"

    # Dados do servidor
    ws["A1"] = "DADOS DO SERVIDOR"
    ws["A1"].font = Font(bold=True, size=14)

    linha = 3
    for campo, valor in dados_servidor.items():
        ws.cell(linha, 1, campo).font = Font(bold=True)
        ws.cell(linha, 2, valor)
        linha += 1

    # Tabela
    linha_tabela = linha + 1
    cabecalho = ["Código", "Descrição", *COLUNAS_VALORES]

    for coluna, valor in enumerate(cabecalho, start=1):
        celula = ws.cell(linha_tabela, coluna, valor)
        celula.font = Font(bold=True)
        celula.fill = PatternFill("solid", fgColor="D9EAF7")
        celula.alignment = Alignment(horizontal="center")

    for i, rubrica in enumerate(rubricas, start=linha_tabela + 1):
        for coluna, campo in enumerate(cabecalho, start=1):
            ws.cell(i, coluna, rubrica[campo])

    # Formatação dos valores financeiros
    primeira_coluna_valor = 3
    ultima_linha = linha_tabela + len(rubricas)

    for row in ws.iter_rows(
        min_row=linha_tabela + 1,
        max_row=ultima_linha,
        min_col=primeira_coluna_valor,
        max_col=len(cabecalho),
    ):
        for celula in row:
            celula.number_format = '#,##0.00'

    # Ajuste simples das larguras
    larguras = {
        1: 12,
        2: 35,
    }
    for coluna, largura in larguras.items():
        ws.column_dimensions[get_column_letter(coluna)].width = largura

    for coluna in range(3, len(cabecalho) + 1):
        ws.column_dimensions[get_column_letter(coluna)].width = 14

    ws.freeze_panes = f"A{linha_tabela + 1}"
    ws.auto_filter.ref = (
        f"A{linha_tabela}:{get_column_letter(len(cabecalho))}{ultima_linha}"
    )

    caminho_saida.parent.mkdir(parents=True, exist_ok=True)
    wb.save(caminho_saida)


def processar(caminho_pdf: Path, caminho_excel: Path) -> None:
    texto = extrair_texto_pdf(caminho_pdf)
    dados_servidor = extrair_dados_servidor(texto)
    rubricas = extrair_rubricas(texto)

    if not dados_servidor:
        raise ValueError("Não foi possível identificar os dados do servidor.")

    if not rubricas:
        raise ValueError("Nenhuma rubrica foi encontrada no PDF.")

    criar_excel(caminho_excel, dados_servidor, rubricas)

    print(f"Servidor: {dados_servidor.get('Nome', 'não identificado')}")
    print(f"Ano: {dados_servidor.get('Ano', 'não identificado')}")
    print(f"Rubricas encontradas: {len(rubricas)}")
    print(f"Excel criado: {caminho_excel}")


def main():
    if len(sys.argv) != 2:
        print("Uso: python src/main.py entrada.pdf")
        sys.exit(1)

    caminho_pdf = Path(sys.argv[1])

    if not caminho_pdf.exists():
        print(f"Arquivo não encontrado: {caminho_pdf}")
        sys.exit(1)

    if caminho_pdf.suffix.lower() != ".pdf":
        print("O arquivo informado não é um PDF.")
        sys.exit(1)

    caminho_excel = caminho_pdf.with_suffix(".xlsx")

    processar(caminho_pdf, caminho_excel)


if __name__ == "__main__":
    main()
