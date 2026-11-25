
from __future__ import annotations

import unicodedata
from pathlib import Path
from datetime import date, datetime
from collections import defaultdict

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


# ============================================================
# CONFIGURAÇÕES
# ============================================================

ROOT = Path(__file__).resolve().parents[1]

PASTA_DADOS = ROOT / "data" / "unprocessed"
PASTA_CADASTRO = ROOT / "cadastro"
PASTA_PROCESSADOS = ROOT / "data" / "processed"

ARQUIVO_DEPARA = PASTA_CADASTRO / "de_para_mercapromo.xlsx"
ARQUIVO_DEPARA_ALTERNATIVO = PASTA_CADASTRO / "de_para.xlsx"

ARQUIVO_RELATORIO = (
    PASTA_PROCESSADOS / "resumo_exploracao_black_friday.txt"
)

ABA_BLACK_FRIDAY = "CERVEJA"

QUANTIDADE_LINHAS_AMOSTRA = 20
QUANTIDADE_EXEMPLOS_COLUNA = 5


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def normalizar_texto(texto: str) -> str:
    """
    Remove acentos e transforma o texto em minúsculas.
    Facilita a identificação de arquivos e abas.
    """
    texto = str(texto).strip().lower()

    texto = unicodedata.normalize("NFKD", texto)

    texto = "".join(
        caractere
        for caractere in texto
        if not unicodedata.combining(caractere)
    )

    return texto


def valor_legivel(valor) -> str:
    """
    Converte valores do Excel para uma representação legível.
    """

    if valor is None:
        return ""

    if isinstance(valor, (datetime, date)):
        return valor.strftime("%d/%m/%Y")

    if isinstance(valor, float):
        if valor.is_integer():
            return str(int(valor))

    texto = str(valor).replace("\n", " ").strip()

    return texto


def encontrar_arquivo_depara() -> Path | None:
    """
    Localiza o arquivo De/Para utilizado pelo projeto.
    """

    if ARQUIVO_DEPARA.exists():
        return ARQUIVO_DEPARA

    if ARQUIVO_DEPARA_ALTERNATIVO.exists():
        return ARQUIVO_DEPARA_ALTERNATIVO

    # Procura outros arquivos com depara no nome
    # dentro da pasta cadastro.
    if PASTA_CADASTRO.exists():

        candidatos = []

        for arquivo in PASTA_CADASTRO.iterdir():

            if not arquivo.is_file():
                continue

            if arquivo.suffix.lower() not in (".xlsx", ".xlsm"):
                continue

            nome = normalizar_texto(arquivo.stem)

            if "depara" in nome or "de_para" in nome:
                candidatos.append(arquivo)

        if candidatos:
            return sorted(candidatos)[0]

    return None


def encontrar_arquivos_black_friday() -> list[Path]:
    """
    Procura arquivos Excel com Black no nome.

    Arquivos de De/Para e arquivos temporários do Excel
    são ignorados.
    """

    if not PASTA_DADOS.exists():
        return []

    arquivos_encontrados = []

    for arquivo in PASTA_DADOS.iterdir():

        if not arquivo.is_file():
            continue

        if arquivo.name.startswith("~$"):
            continue

        if arquivo.suffix.lower() not in (".xlsx", ".xlsm"):
            continue

        nome = normalizar_texto(arquivo.stem)

        if "black" not in nome:
            continue

        if "depara" in nome or "de_para" in nome:
            continue

        arquivos_encontrados.append(arquivo)

    return sorted(arquivos_encontrados)


# ============================================================
# EXPLORAÇÃO DE UMA PLANILHA
# ============================================================

def explorar_planilha(ws, linhas: list[str]) -> None:
    """
    Explora uma aba do Excel:

    - Dimensões
    - Primeiras 20 linhas
    - Colunas
    - Quantidade de valores preenchidos
    - Exemplos de valores encontrados

    Usa as células efetivamente preenchidas para evitar
    percorrer milhares de linhas vazias criadas por formatação.
    """

    linhas.append("")
    linhas.append("=" * 90)
    linhas.append(f"ABA: {ws.title}")
    linhas.append("=" * 90)

    # Obtém somente as células com conteúdo.
    # O acesso a _cells evita percorrer grandes intervalos
    # vazios que podem existir por causa da formatação do Excel.
    dados = {}

    for (numero_linha, numero_coluna), celula in ws._cells.items():

        valor = celula.value

        if valor is None:
            continue

        if isinstance(valor, str) and not valor.strip():
            continue

        dados[(numero_linha, numero_coluna)] = valor

    if not dados:
        linhas.append("A aba não possui dados preenchidos.")
        linhas.append(
            f"Dimensão registrada pelo Excel: "
            f"{ws.max_row} linhas x {ws.max_column} colunas"
        )
        return

    ultima_linha = max(linha for linha, coluna in dados)
    ultima_coluna = max(coluna for linha, coluna in dados)

    linhas.append(
        f"Dimensão com conteúdo: "
        f"{ultima_linha} linhas x {ultima_coluna} colunas"
    )

    linhas.append(
        f"Dimensão registrada pelo Excel: "
        f"{ws.max_row} linhas x {ws.max_column} colunas"
    )

    # --------------------------------------------------------
    # PRIMEIRAS LINHAS
    # --------------------------------------------------------

    linhas.append("")
    linhas.append("-" * 90)
    linhas.append(
        f"AMOSTRA DAS PRIMEIRAS {QUANTIDADE_LINHAS_AMOSTRA} LINHAS"
    )
    linhas.append("-" * 90)

    limite_linhas = min(
        ultima_linha,
        QUANTIDADE_LINHAS_AMOSTRA
    )

    for numero_linha in range(1, limite_linhas + 1):

        valores_linha = []

        for numero_coluna in range(1, ultima_coluna + 1):

            valor = dados.get((numero_linha, numero_coluna))

            if valor is None:
                continue

            letra = get_column_letter(numero_coluna)

            valor_formatado = valor_legivel(valor)

            valores_linha.append(
                f"{letra}={valor_formatado}"
            )

        if valores_linha:

            linhas.append(
                f"Linha {numero_linha}: "
                + " | ".join(valores_linha)
            )

    # --------------------------------------------------------
    # ANÁLISE DAS COLUNAS
    # --------------------------------------------------------

    linhas.append("")
    linhas.append("-" * 90)
    linhas.append("ANÁLISE DAS COLUNAS")
    linhas.append("-" * 90)

    valores_por_coluna = defaultdict(list)

    for (numero_linha, numero_coluna), valor in dados.items():

        valores_por_coluna[numero_coluna].append(
            (numero_linha, valor)
        )

    for numero_coluna in range(1, ultima_coluna + 1):

        valores = valores_por_coluna.get(numero_coluna, [])

        if not valores:
            continue

        letra = get_column_letter(numero_coluna)

        valor_linha_1 = dados.get((1, numero_coluna))

        cabecalho = valor_legivel(valor_linha_1)

        quantidade = len(valores)

        exemplos = []
        exemplos_identificados = set()

        for numero_linha, valor in sorted(valores):

            exemplo = valor_legivel(valor)

            if not exemplo:
                continue

            if exemplo in exemplos_identificados:
                continue

            exemplos.append(exemplo)
            exemplos_identificados.add(exemplo)

            if len(exemplos) >= QUANTIDADE_EXEMPLOS_COLUNA:
                break

        linhas.append("")

        if cabecalho:
            linhas.append(
                f"Coluna {letra} | Linha 1: {cabecalho}"
            )
        else:
            linhas.append(
                f"Coluna {letra} | Sem cabeçalho na linha 1"
            )

        linhas.append(
            f"  Valores preenchidos: {quantidade}"
        )

        if exemplos:
            linhas.append(
                "  Exemplos: " + " | ".join(exemplos)
            )


# ============================================================
# EXPLORAÇÃO DO ARQUIVO BLACK FRIDAY
# ============================================================

def explorar_arquivo_black_friday(
    arquivo: Path,
    linhas: list[str]
) -> None:
    """
    Explora somente a aba CERVEJA do arquivo Black Friday.

    As demais abas são identificadas no relatório,
    mas seus conteúdos não são explorados.
    """

    linhas.append("")
    linhas.append("#" * 100)
    linhas.append("ARQUIVO BLACK FRIDAY")
    linhas.append("#" * 100)

    linhas.append(f"Arquivo: {arquivo.name}")
    linhas.append(f"Caminho: {arquivo}")

    try:

        wb = load_workbook(
            arquivo,
            data_only=True,
            read_only=False
        )

        linhas.append("")
        linhas.append("Abas encontradas no arquivo:")

        for nome_aba in wb.sheetnames:
            linhas.append(f"  - {nome_aba}")

        # Localiza a aba CERVEJA, ignorando diferenças
        # de maiúsculas, minúsculas e espaços.
        aba_cerveja = None

        for ws in wb.worksheets:

            if normalizar_texto(ws.title) == normalizar_texto(
                ABA_BLACK_FRIDAY
            ):
                aba_cerveja = ws
                break

        if aba_cerveja is None:

            linhas.append("")
            linhas.append(
                f"ERRO: A aba '{ABA_BLACK_FRIDAY}' não foi encontrada."
            )

            linhas.append(
                "Nenhuma outra aba do arquivo foi explorada."
            )

        else:

            linhas.append("")
            linhas.append(
                f"Aba selecionada para exploração: {aba_cerveja.title}"
            )

            explorar_planilha(
                aba_cerveja,
                linhas
            )

        wb.close()

    except Exception as erro:

        linhas.append("")
        linhas.append(
            f"ERRO AO ABRIR O ARQUIVO: {erro}"
        )


# ============================================================
# EXPLORAÇÃO DO DE/PARA
# ============================================================

def explorar_arquivo_depara(
    arquivo: Path,
    linhas: list[str]
) -> None:
    """
    Explora o arquivo De/Para utilizado pelo projeto.

    Todas as abas desse arquivo são consideradas.
    """

    linhas.append("")
    linhas.append("#" * 100)
    linhas.append("ARQUIVO DE/PARA")
    linhas.append("#" * 100)

    linhas.append(f"Arquivo: {arquivo.name}")
    linhas.append(f"Caminho: {arquivo}")

    try:

        wb = load_workbook(
            arquivo,
            data_only=True,
            read_only=False
        )

        linhas.append("")
        linhas.append("Abas encontradas no De/Para:")

        for nome_aba in wb.sheetnames:
            linhas.append(f"  - {nome_aba}")

        for ws in wb.worksheets:

            explorar_planilha(
                ws,
                linhas
            )

        wb.close()

    except Exception as erro:

        linhas.append("")
        linhas.append(
            f"ERRO AO ABRIR O DE/PARA: {erro}"
        )


# ============================================================
# EXECUÇÃO PRINCIPAL
# ============================================================

def main() -> None:

    linhas = []

    linhas.append("=" * 100)
    linhas.append("RELATÓRIO DE EXPLORAÇÃO - BLACK FRIDAY")
    linhas.append("=" * 100)

    linhas.append("")
    linhas.append(f"Diretório do projeto: {ROOT}")
    linhas.append(f"Diretório de dados: {PASTA_DADOS}")

    # --------------------------------------------------------
    # LOCALIZAR BLACK FRIDAY
    # --------------------------------------------------------

    arquivos_black_friday = encontrar_arquivos_black_friday()

    linhas.append("")
    linhas.append("=" * 100)
    linhas.append("ARQUIVOS BLACK FRIDAY ENCONTRADOS")
    linhas.append("=" * 100)

    if not arquivos_black_friday:

        linhas.append(
            "Nenhum arquivo Excel com 'Black' no nome foi encontrado."
        )

    else:

        for arquivo in arquivos_black_friday:

            linhas.append(f"  - {arquivo.name}")

        # Explora todos os arquivos encontrados,
        # mas somente a aba CERVEJA de cada um.
        for arquivo in arquivos_black_friday:

            explorar_arquivo_black_friday(
                arquivo,
                linhas
            )

    # --------------------------------------------------------
    # LOCALIZAR E EXPLORAR DE/PARA
    # --------------------------------------------------------

    arquivo_depara = encontrar_arquivo_depara()

    if arquivo_depara is None:

        linhas.append("")
        linhas.append("=" * 100)
        linhas.append("DE/PARA")
        linhas.append("=" * 100)

        linhas.append(
            "ERRO: Não foi possível localizar o arquivo De/Para."
        )

        linhas.append(
            f"Verifique a pasta: {PASTA_CADASTRO}"
        )

    else:

        explorar_arquivo_depara(
            arquivo_depara,
            linhas
        )

    # --------------------------------------------------------
    # SALVAR RELATÓRIO
    # --------------------------------------------------------

    PASTA_PROCESSADOS.mkdir(
        parents=True,
        exist_ok=True
    )

    conteudo = "\n".join(linhas)

    ARQUIVO_RELATORIO.write_text(
        conteudo,
        encoding="utf-8"
    )

    print("")
    print("=" * 70)
    print("EXPLORAÇÃO CONCLUÍDA")
    print("=" * 70)

    print(
        f"Relatório salvo em:\n{ARQUIVO_RELATORIO}"
    )

    print("")
    print(
        "Arquivo(s) Black Friday encontrado(s): "
        f"{len(arquivos_black_friday)}"
    )

    print(
        "Aba explorada no Black Friday: "
        f"{ABA_BLACK_FRIDAY}"
    )

    print("")
    print("Abra o arquivo TXT para conferir os resultados.")


if __name__ == "__main__":
    main()
