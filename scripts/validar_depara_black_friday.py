
from __future__ import annotations

import csv
import json
import unicodedata

from collections import Counter, defaultdict
from difflib import SequenceMatcher
from pathlib import Path

from openpyxl import load_workbook


# ============================================================
# CONFIGURAÇÕES
# ============================================================

ROOT = Path(__file__).resolve().parents[1]

ARQUIVO_BLACK_FRIDAY = (
    ROOT / "data" / "processed" / "black_friday_cerveja.json"
)

ARQUIVO_DEPARA = (
    ROOT / "cadastro" / "de_para_mercapromo.xlsx"
)

PASTA_IMAGENS = ROOT / "images" / "img"

PASTA_SAIDA = ROOT / "data" / "processed"

ARQUIVO_RELATORIO = (
    PASTA_SAIDA / "validacao_depara_black_friday.txt"
)

ARQUIVO_CSV = (
    PASTA_SAIDA / "validacao_depara_black_friday.csv"
)

# Similaridade mínima para apresentar uma sugestão.
# Não significa correspondência confirmada.
MIN_SCORE_SUGESTAO = 0.35

QUANTIDADE_SUGESTOES = 3


# ============================================================
# NORMALIZAÇÃO DE NOMES
# ============================================================

def normalizar(valor) -> str:
    """
    Padroniza caixa, acentos, pontuação e espaços.

    Não remove tamanhos, embalagens ou características
    dos produtos, pois essas diferenças podem representar SKUs
    distintos.
    """

    if valor is None:
        return ""

    texto = str(valor).strip().upper()

    texto = unicodedata.normalize("NFKD", texto)

    texto = "".join(
        caractere
        for caractere in texto
        if not unicodedata.combining(caractere)
    )

    caracteres = []

    for caractere in texto:

        if caractere.isalnum():
            caracteres.append(caractere)
        else:
            caracteres.append(" ")

    return " ".join("".join(caracteres).split())


# ============================================================
# LEITURA DO DE/PARA
# ============================================================

def carregar_depara() -> list[dict]:

    if not ARQUIVO_DEPARA.exists():
        raise FileNotFoundError(
            f"De/Para não encontrado: {ARQUIVO_DEPARA}"
        )

    wb = load_workbook(
        ARQUIVO_DEPARA,
        data_only=True,
        read_only=True
    )

    try:

        ws = wb.active

        cabecalhos = {
            normalizar(cell.value): indice
            for indice, cell in enumerate(
                ws[1],
                start=1
            )
            if cell.value is not None
        }

        campos_obrigatorios = [
            "PROMAX",
            "SKU",
            "NOME IMG",
            "CAMINHO",
            "NOME",
        ]

        faltantes = [
            campo
            for campo in campos_obrigatorios
            if campo not in cabecalhos
        ]

        if faltantes:
            raise ValueError(
                f"Colunas ausentes no De/Para: {faltantes}"
            )

        def obter(linha, nome_campo):

            coluna = cabecalhos[nome_campo]

            valor = ws.cell(
                row=linha,
                column=coluna
            ).value

            if valor is None:
                return ""

            return str(valor).strip()

        produtos = []

        for linha in range(2, ws.max_row + 1):

            nome = obter(linha, "NOME")

            if not nome:
                continue

            produtos.append({

                "nome": nome,

                "nome_normalizado": normalizar(nome),

                "promax": obter(linha, "PROMAX"),

                "sku": obter(linha, "SKU"),

                "nome_img": obter(linha, "NOME IMG"),

                "caminho": obter(linha, "CAMINHO"),

                "linha_depara": linha,

            })

        return produtos

    finally:
        wb.close()


# ============================================================
# INDEXAÇÃO DAS IMAGENS
# ============================================================

def indexar_imagens() -> dict[str, Path]:

    imagens = {}

    if not PASTA_IMAGENS.exists():
        return imagens

    extensoes = {
        ".png",
        ".jpg",
        ".jpeg",
        ".webp",
    }

    for arquivo in PASTA_IMAGENS.rglob("*"):

        if not arquivo.is_file():
            continue

        if arquivo.suffix.lower() not in extensoes:
            continue

        imagens[arquivo.name.lower()] = arquivo

    return imagens


def localizar_imagem(
    caminho: str,
    indice_imagens: dict[str, Path]
) -> Path | None:

    if not caminho:
        return None

    return indice_imagens.get(
        Path(caminho).name.lower()
    )


# ============================================================
# SUGESTÕES DE PRODUTOS
# ============================================================

def encontrar_sugestoes(
    produto: str,
    produtos_depara: list[dict]
) -> list[dict]:

    nome_normalizado = normalizar(produto)

    candidatos = []

    for item in produtos_depara:

        score = SequenceMatcher(
            None,
            nome_normalizado,
            item["nome_normalizado"]
        ).ratio()

        if score >= MIN_SCORE_SUGESTAO:

            candidatos.append({
                **item,
                "similaridade": round(score * 100, 1)
            })

    candidatos.sort(
        key=lambda item: item["similaridade"],
        reverse=True
    )

    return candidatos[:QUANTIDADE_SUGESTOES]


# ============================================================
# VALIDAÇÃO DE UM PRODUTO
# ============================================================

def validar_produto(
    produto: str,
    quantidade_ofertas: int,
    produtos_depara: list[dict],
    indice_imagens: dict[str, Path]
) -> dict:

    nome_normalizado = normalizar(produto)

    correspondencias = [
        item
        for item in produtos_depara
        if item["nome_normalizado"] == nome_normalizado
    ]

    resultado = {
        "produto": produto,
        "quantidade_ofertas": quantidade_ofertas,
        "status": "",
        "sku": "",
        "promax": "",
        "nome_img": "",
        "caminho": "",
        "imagem_encontrada": "",
        "sugestoes": [],
    }

    # --------------------------------------------------------
    # CORRESPONDÊNCIA EXATA
    # --------------------------------------------------------

    if len(correspondencias) == 1:

        item = correspondencias[0]

        imagem = localizar_imagem(
            item["caminho"],
            indice_imagens
        )

        resultado.update({
            "status": "ENCONTRADO",
            "sku": item["sku"],
            "promax": item["promax"],
            "nome_img": item["nome_img"],
            "caminho": item["caminho"],
            "imagem_encontrada": (
                "SIM" if imagem else "NÃO"
            ),
        })

        return resultado

    # --------------------------------------------------------
    # MAIS DE UMA CORRESPONDÊNCIA
    # --------------------------------------------------------

    if len(correspondencias) > 1:

        resultado["status"] = "DUPLICADO_NO_DEPARA"

        resultado["sugestoes"] = [
            {
                "nome": item["nome"],
                "sku": item["sku"],
                "similaridade": 100.0,
            }
            for item in correspondencias
        ]

        return resultado

    # --------------------------------------------------------
    # SEM CORRESPONDÊNCIA EXATA
    # --------------------------------------------------------

    sugestoes = encontrar_sugestoes(
        produto,
        produtos_depara
    )

    resultado["sugestoes"] = [
        {
            "nome": item["nome"],
            "sku": item["sku"],
            "similaridade": item["similaridade"],
        }
        for item in sugestoes
    ]

    if sugestoes:
        resultado["status"] = "REVISAR_SUGESTAO"
    else:
        resultado["status"] = "SEM_CORRESPONDENCIA"

    return resultado


# ============================================================
# EXECUTAR VALIDAÇÃO
# ============================================================

def main():

    print("Iniciando validação do De/Para...")

    if not ARQUIVO_BLACK_FRIDAY.exists():
        raise FileNotFoundError(
            f"Base Black Friday não encontrada: "
            f"{ARQUIVO_BLACK_FRIDAY}"
        )

    # Carregar ofertas extraídas.
    with ARQUIVO_BLACK_FRIDAY.open(
        "r",
        encoding="utf-8"
    ) as arquivo:

        ofertas = json.load(arquivo)

    # Quantidade de ofertas por produto.
    contagem_produtos = Counter(
        oferta["PRODUTO"]
        for oferta in ofertas
    )

    # Carregar De/Para.
    produtos_depara = carregar_depara()

    print(
        f"Produtos cadastrados no De/Para: "
        f"{len(produtos_depara)}"
    )

    # Indexar imagens.
    indice_imagens = indexar_imagens()

    print(
        f"Imagens encontradas na pasta: "
        f"{len(indice_imagens)}"
    )

    # Validar cada produto distinto.
    resultados = []

    for produto, quantidade in contagem_produtos.items():

        resultado = validar_produto(
            produto,
            quantidade,
            produtos_depara,
            indice_imagens
        )

        resultados.append(resultado)

    # --------------------------------------------------------
    # CONTADORES
    # --------------------------------------------------------

    contagem_status = Counter(
        item["status"]
        for item in resultados
    )

    total_encontrados = contagem_status["ENCONTRADO"]

    total_revisar = contagem_status["REVISAR_SUGESTAO"]

    total_duplicados = contagem_status["DUPLICADO_NO_DEPARA"]

    total_sem_correspondencia = contagem_status[
        "SEM_CORRESPONDENCIA"
    ]

    total_imagens = sum(
        1
        for item in resultados
        if item["imagem_encontrada"] == "SIM"
    )

    total_ofertas_encontradas = sum(
        item["quantidade_ofertas"]
        for item in resultados
        if item["status"] == "ENCONTRADO"
    )

    # --------------------------------------------------------
    # RELATÓRIO TXT
    # --------------------------------------------------------

    linhas = []

    linhas.append("=" * 80)
    linhas.append("VALIDAÇÃO DE/PARA - BLACK FRIDAY")
    linhas.append("=" * 80)

    linhas.append("")
    linhas.append(f"Total de ofertas: {len(ofertas)}")
    linhas.append(
        f"Produtos distintos: {len(resultados)}"
    )

    linhas.append("")
    linhas.append("RESUMO")

    linhas.append(
        f"Correspondências exatas: {total_encontrados}"
    )

    linhas.append(
        f"Ofertas com correspondência exata: "
        f"{total_ofertas_encontradas}"
    )

    linhas.append(
        f"Produtos com imagem localizada: {total_imagens}"
    )

    linhas.append(
        f"Precisam de revisão: {total_revisar}"
    )

    linhas.append(
        f"Duplicados no De/Para: {total_duplicados}"
    )

    linhas.append(
        f"Sem correspondência: {total_sem_correspondencia}"
    )

    # --------------------------------------------------------
    # DETALHAMENTO
    # --------------------------------------------------------

    for resultado in resultados:

        linhas.append("")
        linhas.append("-" * 80)

        linhas.append(
            f"PRODUTO: {resultado['produto']}"
        )

        linhas.append(
            f"QUANTIDADE DE OFERTAS: "
            f"{resultado['quantidade_ofertas']}"
        )

        linhas.append(
            f"STATUS: {resultado['status']}"
        )

        if resultado["status"] == "ENCONTRADO":

            linhas.append(
                f"SKU: {resultado['sku']}"
            )

            linhas.append(
                f"PROMAX: {resultado['promax']}"
            )

            linhas.append(
                f"NOME IMG: {resultado['nome_img']}"
            )

            linhas.append(
                f"CAMINHO: {resultado['caminho']}"
            )

            linhas.append(
                f"IMAGEM LOCALIZADA: "
                f"{resultado['imagem_encontrada']}"
            )

        elif resultado["sugestoes"]:

            linhas.append("SUGESTÕES PARA REVISÃO:")

            for sugestao in resultado["sugestoes"]:

                linhas.append(
                    f"  - {sugestao['nome']} | "
                    f"SKU: {sugestao['sku']} | "
                    f"Similaridade: "
                    f"{sugestao['similaridade']}%"
                )

    PASTA_SAIDA.mkdir(
        parents=True,
        exist_ok=True
    )

    ARQUIVO_RELATORIO.write_text(
        "\n".join(linhas),
        encoding="utf-8"
    )

    # --------------------------------------------------------
    # CSV PARA CONFERÊNCIA LOCAL
    # --------------------------------------------------------

    with ARQUIVO_CSV.open(
        "w",
        encoding="utf-8-sig",
        newline=""
    ) as arquivo:

        campos = [
            "produto",
            "quantidade_ofertas",
            "status",
            "sku",
            "promax",
            "nome_img",
            "caminho",
            "imagem_encontrada",
            "sugestoes",
        ]

        escritor = csv.DictWriter(
            arquivo,
            fieldnames=campos,
            delimiter=";"
        )

        escritor.writeheader()

        for resultado in resultados:

            linha = resultado.copy()

            linha["sugestoes"] = " | ".join(
                (
                    f"{s['nome']} "
                    f"(SKU {s['sku']}, "
                    f"{s['similaridade']}%)"
                )
                for s in resultado["sugestoes"]
            )

            escritor.writerow(linha)

    # --------------------------------------------------------
    # RESULTADO NO TERMINAL
    # --------------------------------------------------------

    print("")
    print("=" * 60)
    print("VALIDAÇÃO CONCLUÍDA")
    print("=" * 60)

    print(f"Ofertas: {len(ofertas)}")
    print(f"Produtos distintos: {len(resultados)}")
    print(f"Correspondências exatas: {total_encontrados}")
    print(f"Precisam de revisão: {total_revisar}")
    print(f"Duplicados: {total_duplicados}")
    print(f"Sem correspondência: {total_sem_correspondencia}")
    print(f"Imagens localizadas: {total_imagens}")

    print("")
    print(f"Relatório: {ARQUIVO_RELATORIO}")
    print(f"CSV: {ARQUIVO_CSV}")


if __name__ == "__main__":
    main()
