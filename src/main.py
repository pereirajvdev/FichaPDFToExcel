import re
import sys
from pathlib import Path

import pdfplumber
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# ============================================================
# CONFIGURAÇÕES
# ============================================================

COLUNAS_VALORES = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro",
    "13º salário", "Rescisão", "Total",
]

MESES = COLUNAS_VALORES[:12]

PADRAO_DINHEIRO = r"-?\d{1,3}(?:\.\d{3})*,\d{2}"

# Nome do arquivo usado como modelo
ARQUIVO_TEMPLATE = Path(
    r"C:\Users\Joao Castro\Desktop\JOAO\Anexo V\Servidores\AnexoVMacro.xlsm"
)
# Nome da aba do template
NOME_ABA = "Plan1"


# ============================================================
# PDF
# ============================================================

def extrair_texto_pdf(caminho_pdf: Path) -> str:
    """Extrai o texto de todas as páginas do PDF."""

    paginas = []

    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text() or ""
            paginas.append(texto)

    return "\n".join(paginas)


# ============================================================
# VALORES
# ============================================================

def normalizar_valor(valor: str) -> float:
    """Converte valor brasileiro, como 1.234,56, para float."""

    return float(
        valor.replace(".", "").replace(",", ".")
    )


# ============================================================
# DADOS DO SERVIDOR
# ============================================================

def extrair_dados_servidor(texto: str) -> dict:
    """Extrai os dados do servidor do cabeçalho da ficha."""

    dados = {}

    padrao_servidor = re.search(
        r"(?m)^(\d{6})\s+(.+?)\s+Secretaria:",
        texto
    )

    if padrao_servidor:
        dados["Matrícula"] = padrao_servidor.group(1)
        dados["Nome"] = padrao_servidor.group(2).strip()

    padrao = re.search(
        r"Ficha Financeira Resumo Geral do Ano de (\d{4})",
        texto
    )

    if padrao:
        dados["Ano"] = int(padrao.group(1))

    padrao = re.search(
        r"Secretaria:\s*(.*?)\s+Conta:",
        texto
    )

    if padrao:
        dados["Secretaria"] = padrao.group(1).strip()

    padrao = re.search(
        r"Conta:\s*(\S+)\s+Ano:",
        texto
    )

    if padrao:
        dados["Conta"] = padrao.group(1)

    padrao = re.search(
        r"Cargo:\s*(.*?)\s+CTPS:",
        texto
    )

    if padrao:
        dados["Cargo"] = padrao.group(1).strip()

    padrao = re.search(
        r"CTPS:\s*(\S+)\s+Pis/Pasep:",
        texto
    )

    if padrao:
        dados["CTPS"] = padrao.group(1)

    padrao = re.search(
        r"Pis/Pasep:\s*(\S+)\s+CPF:",
        texto
    )

    if padrao:
        dados["PIS/PASEP"] = padrao.group(1)

    padrao = re.search(
        r"CPF:\s*(\S+)\s+Dt\. Nasc\.:",
        texto
    )

    if padrao:
        dados["CPF"] = padrao.group(1)

    padrao = re.search(
        r"Dt\. Nasc\.\s*:\s*(\S+)\s+Dt\. Admissão:\s*(\S+)",
        texto
    )

    if padrao:
        dados["Data de nascimento"] = padrao.group(1)
        dados["Data de admissão"] = padrao.group(2)

    return dados


# ============================================================
# RUBRICAS
# ============================================================

def encontrar_divisor(texto: str) -> int:
    """
    Encontra a posição de 'Total de Vencimentos'.

    Tudo que aparece antes do divisor será considerado
    vencimento e tudo depois será considerado desconto.
    """

    padrao = re.search(
        r"Total\s+de\s+Vencimentos",
        texto,
        flags=re.IGNORECASE
    )

    if not padrao:
        raise ValueError(
            "Não foi possível encontrar 'Total de Vencimentos' no PDF."
        )

    return padrao.start()


def extrair_rubricas(texto: str) -> tuple[list[dict], list[dict]]:
    """
    Extrai as rubricas e separa em Vencimentos e Descontos.
    """

    rubricas_vencimentos = []
    rubricas_descontos = []

    padrao_valor = re.compile(PADRAO_DINHEIRO)

    divisor = encontrar_divisor(texto)

    for linha in texto.splitlines():

        linha = linha.strip()

        if not re.match(r"^\d{5}\s+", linha):
            continue

        codigo = linha[:5]
        restante = linha[5:].strip()

        valores_encontrados = list(
            padrao_valor.finditer(restante)
        )

        if len(valores_encontrados) != 15:
            print(
                "Aviso: rubrica ignorada por quantidade "
                f"inesperada de valores: {linha}"
            )
            continue

        descricao = restante[
            :valores_encontrados[0].start()
        ].strip()

        valores = [
            normalizar_valor(m.group())
            for m in valores_encontrados
        ]

        rubrica = {
            "Código": codigo,
            "Descrição": descricao,
            **dict(zip(COLUNAS_VALORES, valores)),
        }

        posicao_linha = texto.find(linha)

        if posicao_linha < divisor:
            rubricas_vencimentos.append(rubrica)
        else:
            rubricas_descontos.append(rubrica)

    return rubricas_vencimentos, rubricas_descontos


# ============================================================
# LOCALIZAR FIM DO TEMPLATE
# ============================================================

def encontrar_ultima_linha_preenchida(ws) -> int:
    """
    Encontra a última linha que realmente possui conteúdo
    no template.

    Isso evita depender de uma linha fixa como 33.
    """

    ultima_linha = 0

    for row in ws.iter_rows():

        for celula in row:

            if celula.value is not None:
                ultima_linha = max(
                    ultima_linha,
                    celula.row
                )

    return ultima_linha


# ============================================================
# ESCREVER TABELA
# ============================================================

def escrever_tabela(
    ws,
    linha_inicial: int,
    titulo: str,
    rubricas: list[dict],
) -> int:
    """Escreve uma tabela de rubricas e retorna a próxima linha."""

    # --------------------------------------------------------
    # TÍTULO
    # --------------------------------------------------------

    ws.cell(
        linha_inicial,
        1,
        titulo
    )

    ws.cell(
        linha_inicial,
        1
    ).font = Font(
        bold=True,
        size=13
    )

    linha_tabela = linha_inicial + 1

    # --------------------------------------------------------
    # CABEÇALHO
    # --------------------------------------------------------

    cabecalho = [
        "Código",
        "Descrição",
        *COLUNAS_VALORES,
    ]

    for coluna, valor in enumerate(
        cabecalho,
        start=1
    ):

        celula = ws.cell(
            linha_tabela,
            coluna,
            valor
        )

        celula.font = Font(
            bold=True
        )

        celula.fill = PatternFill(
            "solid",
            fgColor="D9EAF7"
        )

        celula.alignment = Alignment(
            horizontal="center"
        )

    # --------------------------------------------------------
    # RUBRICAS
    # --------------------------------------------------------

    for i, rubrica in enumerate(
        rubricas,
        start=linha_tabela + 1
    ):

        for coluna, campo in enumerate(
            cabecalho,
            start=1
        ):

            ws.cell(
                i,
                coluna,
                rubrica[campo]
            )

    # --------------------------------------------------------
    # FORMATAÇÃO FINANCEIRA
    # --------------------------------------------------------

    primeira_coluna_valor = 3

    ultima_linha = (
        linha_tabela + len(rubricas)
    )

    for row in ws.iter_rows(
        min_row=linha_tabela + 1,
        max_row=ultima_linha,
        min_col=primeira_coluna_valor,
        max_col=len(cabecalho),
    ):

        for celula in row:
            celula.number_format = '#,##0.00'

    # --------------------------------------------------------
    # PRÓXIMA LINHA
    # --------------------------------------------------------

    return ultima_linha + 2


# ============================================================
# INSERIR DADOS NO TEMPLATE
# ============================================================

def inserir_dados_no_template(
    caminho_template: Path,
    caminho_saida: Path,
    dados_servidor: dict,
    rubricas_vencimentos: list[dict],
    rubricas_descontos: list[dict],
) -> None:

    # --------------------------------------------------------
    # ABRE O TEMPLATE
    # --------------------------------------------------------

    wb = load_workbook(
        caminho_template,
        keep_vba=True,
        data_only=False
    )

    # --------------------------------------------------------
    # SELECIONA A PLANILHA
    # --------------------------------------------------------

    if NOME_ABA not in wb.sheetnames:
        raise ValueError(
            f"A aba '{NOME_ABA}' não foi encontrada no template."
        )

    ws = wb[NOME_ABA]

    # --------------------------------------------------------
    # ENCONTRA A ÚLTIMA LINHA DO ANEXO V
    # --------------------------------------------------------

    ultima_linha = encontrar_ultima_linha_preenchida(ws)

    # Deixa uma linha em branco entre o template
    # e os dados da ficha financeira.

    linha = ultima_linha + 2

    # --------------------------------------------------------
    # DADOS DO SERVIDOR
    # --------------------------------------------------------

    ws.cell(
        linha,
        1,
        "DADOS DO SERVIDOR"
    )

    ws.cell(
        linha,
        1
    ).font = Font(
        bold=True,
        size=14
    )

    linha += 2

    for campo, valor in dados_servidor.items():

        ws.cell(
            linha,
            1,
            campo
        ).font = Font(
            bold=True
        )

        ws.cell(
            linha,
            2,
            valor
        )

        linha += 1

    # --------------------------------------------------------
    # VENCIMENTOS
    # --------------------------------------------------------

    linha = escrever_tabela(
        ws,
        linha + 1,
        "VENCIMENTOS",
        rubricas_vencimentos,
    )

    # --------------------------------------------------------
    # DESCONTOS
    # --------------------------------------------------------

    linha = escrever_tabela(
        ws,
        linha,
        "DESCONTOS",
        rubricas_descontos,
    )

    # --------------------------------------------------------
    # RECÁLCULO DAS FÓRMULAS
    # --------------------------------------------------------

    # O template continua contendo suas fórmulas.
    # Esta configuração solicita ao Excel que recalcule
    # as fórmulas quando o arquivo for aberto.

    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except AttributeError:
        pass

    # --------------------------------------------------------
    # SALVAR NOVO ARQUIVO
    # --------------------------------------------------------

    caminho_saida.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    wb.save(caminho_saida)


# ============================================================
# PROCESSAMENTO
# ============================================================

def processar(
    caminho_pdf: Path,
    caminho_template: Path,
    caminho_saida: Path,
) -> None:

    print("Lendo PDF...")

    texto = extrair_texto_pdf(
        caminho_pdf
    )

    dados_servidor = extrair_dados_servidor(
        texto
    )

    (
        rubricas_vencimentos,
        rubricas_descontos,
    ) = extrair_rubricas(texto)

    if not dados_servidor:

        raise ValueError(
            "Não foi possível identificar "
            "os dados do servidor."
        )

    if (
        not rubricas_vencimentos
        and not rubricas_descontos
    ):

        raise ValueError(
            "Nenhuma rubrica foi encontrada no PDF."
        )

    print("Inserindo dados no template...")

    inserir_dados_no_template(
        caminho_template,
        caminho_saida,
        dados_servidor,
        rubricas_vencimentos,
        rubricas_descontos,
    )

    # --------------------------------------------------------
    # INFORMAÇÕES
    # --------------------------------------------------------

    print(
        f"Servidor: "
        f"{dados_servidor.get('Nome', 'não identificado')}"
    )

    print(
        f"Ano: "
        f"{dados_servidor.get('Ano', 'não identificado')}"
    )

    print(
        f"Vencimentos: "
        f"{len(rubricas_vencimentos)} rubricas"
    )

    print(
        f"Descontos: "
        f"{len(rubricas_descontos)} rubricas"
    )

    print(
        f"Total de rubricas: "
        f"{len(rubricas_vencimentos) + len(rubricas_descontos)}"
    )

    print(
        f"Excel criado: {caminho_saida}"
    )


# ============================================================
# MAIN
# ============================================================

def main():

    if len(sys.argv) != 2:

        print(
            "Uso: python src/main.py entrada.pdf"
        )

        sys.exit(1)

    caminho_pdf = Path(
        sys.argv[1]
    )

    if not caminho_pdf.exists():

        print(
            f"Arquivo não encontrado: {caminho_pdf}"
        )

        sys.exit(1)

    if caminho_pdf.suffix.lower() != ".pdf":

        print(
            "O arquivo informado não é um PDF."
        )

        sys.exit(1)

    # --------------------------------------------------------
    # TEMPLATE
    # --------------------------------------------------------

    if not ARQUIVO_TEMPLATE.exists():

        print(
            f"Template não encontrado: {ARQUIVO_TEMPLATE}"
        )

        print(
            "Coloque o AnexoVMacro.xlsm na pasta "
            "onde o programa está sendo executado."
        )

        sys.exit(1)

    # --------------------------------------------------------
    # SAÍDA
    # --------------------------------------------------------

    # Exemplo:
    #
    # entrada.pdf
    #
    # vira:
    #
    # entrada.xlsm

    caminho_saida = caminho_pdf.with_suffix(
        ".xlsm"
    )

    # --------------------------------------------------------
    # PROCESSAR
    # --------------------------------------------------------

    try:

        processar(
            caminho_pdf,
            ARQUIVO_TEMPLATE,
            caminho_saida,
        )

    except Exception as erro:

        print(
            f"Erro: {erro}"
        )

        sys.exit(1)


if __name__ == "__main__":
    main()