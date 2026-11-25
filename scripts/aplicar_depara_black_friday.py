
from __future__ import annotations

import json
import unicodedata

from pathlib import Path
from collections import Counter

from openpyxl import load_workbook


# ============================================================
# CAMINHOS
# ============================================================

ROOT = Path(__file__).resolve().parents[1]

ARQUIVO_BASE = (
    ROOT / "data" / "processed" / "black_friday_cerveja.json"
)

ARQUIVO_DEPARA = (
    ROOT / "cadastro" / "de_para_mercapromo.xlsx"
)

PASTA_IMAGENS = ROOT / "images" / "img"

ARQUIVO_SAIDA = (
    ROOT / "data" / "processed"
    / "black_friday_cerveja_enriquecido.json"
)


# ============================================================
# MAPEAMENTO MANUAL APROVADO
# PRODUTO BLACK FRIDAY -> NOME NO DE/PARA
# ============================================================

MAPEAMENTO = {

    "Corona LN 330ml":
        "CORONA EXTRA N LONG NECK 330ML CX C/24 NPAL",

    "Stella Pure Gold LN 330ml":
        "STELLA ARTOIS PURE GOLD LT SLEEK 350ML C 8 CX CARTAO",

    "Michelob LT 269ml":
        "MICHELOB ULTRA N LT 269ML CX CARTAO C/8 NPAL",

    "Original LT 269ml - Multipack 15un":
        "ORIGINAL LT 269ML SH C 15 MULTPACK",

    "Skol Pilsen LT 269ml - Multipack 15un":
        "SKOL LT 269ML SH C/15 MULTPACK 15",

    "Michelob LN 330ml":
        "MICHELOB ULTRA N LONG NECK 330ML SIX-PACK SHRINK C/4",

    "Original LT 269ml":
        "ORIGINAL LT 269ML SH C15 NPAL",

    "Skol Pilsen LT 269ml":
        "SKOL LT 269ML SH C15 NPAL",

    "Brahma Chopp LT 269ml":
        "BRAHMA CHOPP LT 269ML SH C15 NPAL",

    "Michelob LT 350ml":
        "MICHELOB ULTRA N LT SLEEK 350ML C 8 CX CARTAO",

    "Stella Pure Gold LT 350ml":
        "STELLA ARTOIS PURE GOLD LT SLEEK 350ML C 8 CX CARTAO",

    "Antarctica Pilsen LT 269ml":
        "ANTARCTICA PILSEN LT 269ML SH TRANSP C/15 NPAL",
}


# ============================================================
# NORMALIZAÇÃO
# ============================================================

def normalizar(valor) -> str:

    texto = str(valor or "").strip().upper()

    texto = unicodedata.normalize("NFKD", texto)

    texto = "".join(
        c for c in texto
        if not unicodedata.combining(c)
    )

    return " ".join(
        "".join(
            c if c.isalnum() else " "
            for c in texto
        ).split()
    )


# ============================================================
# CARREGAR DE/PARA
# ============================================================

def carregar_depara() -> dict:

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

        obrigatorias = [
            "NOME",
            "SKU",
            "PROMAX",
            "NOME IMG",
            "CAMINHO",
        ]

        for campo in obrigatorias:

            if campo not in cabecalhos:
                raise ValueError(
                    f"Coluna ausente no De/Para: {campo}"
                )

        indice_nome = cabecalhos["NOME"] - 1

        indices = {
            campo: indice - 1
            for campo, indice in cabecalhos.items()
        }

        produtos = {}

        for linha in ws.iter_rows(
            min_row=2,
            values_only=True
        ):

            nome = linha[indice_nome]

            if not nome:
                continue

            nome_normalizado = normalizar(nome)

            registro = {
                "nome": str(nome).strip(),
                "sku": str(
                    linha[indices["SKU"]] or ""
                ).strip(),
                "promax": str(
                    linha[indices["PROMAX"]] or ""
                ).strip(),
                "nome_img": str(
                    linha[indices["NOME IMG"]] or ""
                ).strip(),
                "caminho": str(
                    linha[indices["CAMINHO"]] or ""
                ).strip(),
            }

            produtos.setdefault(
                nome_normalizado, []
            ).append(registro)

        return produtos

    finally:
        wb.close()


# ============================================================
# LOCALIZAR IMAGENS
# ============================================================

def indexar_imagens() -> dict:

    imagens = {}

    if not PASTA_IMAGENS.exists():
        return imagens

    for arquivo in PASTA_IMAGENS.rglob("*"):

        if not arquivo.is_file():
            continue

        if arquivo.suffix.lower() not in {
            ".png", ".jpg", ".jpeg", ".webp"
        }:
            continue

        imagens.setdefault(
            arquivo.name.lower(),
            arquivo
        )

    return imagens


# ============================================================
# LOCALIZAR CORRESPONDÊNCIA
# ============================================================

def resolver_produto(
    produto: str,
    produtos_depara: dict
) -> dict:

    nome_depara = MAPEAMENTO.get(
        produto,
        produto
    )

    chave = normalizar(nome_depara)

    correspondencias = produtos_depara.get(
        chave, []
    )

    if len(correspondencias) == 0:

        raise ValueError(
            f"Não encontrado no De/Para: "
            f"{produto} -> {nome_depara}"
        )

    if len(correspondencias) > 1:

        raise ValueError(
            f"Correspondência duplicada no De/Para: "
            f"{nome_depara}"
        )

    return correspondencias[0]


# ============================================================
# EXECUÇÃO
# ============================================================

def main():

    print("Iniciando aplicação do De/Para...")

    if not ARQUIVO_BASE.exists():
        raise FileNotFoundError(ARQUIVO_BASE)

    if not ARQUIVO_DEPARA.exists():
        raise FileNotFoundError(ARQUIVO_DEPARA)

    with ARQUIVO_BASE.open(
        "r",
        encoding="utf-8"
    ) as arquivo:

        ofertas = json.load(arquivo)

    produtos_depara = carregar_depara()

    imagens = indexar_imagens()

    erros = []

    produtos_resolvidos = {}

    # Resolver uma vez por produto distinto.
    for produto in sorted({
        oferta["PRODUTO"]
        for oferta in ofertas
    }):

        try:

            registro = resolver_produto(
                produto,
                produtos_depara
            )

            caminho = registro["caminho"]

            imagem = imagens.get(
                Path(caminho).name.lower()
            ) if caminho else None

            produtos_resolvidos[produto] = {
                **registro,
                "imagem_existe": imagem is not None,
                "imagem_local": (
                    str(imagem.relative_to(ROOT))
                    if imagem
                    else None
                ),
            }

        except ValueError as erro:

            erros.append(str(erro))

    if erros:

        print("")
        print("ERROS DE MAPEAMENTO:")

        for erro in erros:
            print("-", erro)

        raise RuntimeError(
            "O enriquecimento foi interrompido. "
            "Nenhum JSON final foi gerado."
        )

    # Acrescentar dados do De/Para a cada oferta.
    ofertas_enriquecidas = []

    for oferta in ofertas:

        produto = oferta["PRODUTO"]

        cadastro = produtos_resolvidos[produto]

        nova_oferta = {
            **oferta,
            "DEPARA_NOME": cadastro["nome"],
            "SKU": cadastro["sku"],
            "PROMAX": cadastro["promax"],
            "NOME_IMG": cadastro["nome_img"],
            "CAMINHO_IMAGEM": cadastro["caminho"],
            "IMAGEM_EXISTE": cadastro["imagem_existe"],
            "IMAGEM_LOCAL": cadastro["imagem_local"],
        }

        ofertas_enriquecidas.append(nova_oferta)

    # Salvar somente quando todos os produtos foram resolvidos.
    ARQUIVO_SAIDA.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    ARQUIVO_SAIDA.write_text(
        json.dumps(
            ofertas_enriquecidas,
            ensure_ascii=False,
            indent=2
        ),
        encoding="utf-8"
    )

    # Resumo.
    total_produtos = len(produtos_resolvidos)

    total_imagens = sum(
        1
        for produto in produtos_resolvidos.values()
        if produto["imagem_existe"]
    )

    total_sem_imagem = total_produtos - total_imagens

    distribuicao_skus = Counter(
        produto["sku"]
        for produto in produtos_resolvidos.values()
    )

    print("")
    print("=" * 60)
    print("DE/PARA BLACK FRIDAY APLICADO")
    print("=" * 60)

    print(f"Ofertas enriquecidas: {len(ofertas_enriquecidas)}")
    print(f"Produtos distintos: {total_produtos}")
    print(f"Produtos associados: {total_produtos}")
    print(f"Imagens localizadas: {total_imagens}")
    print(f"Imagens ausentes: {total_sem_imagem}")

    print("")
    print("SKUs associados:")

    for sku, quantidade in sorted(distribuicao_skus.items()):
        print(f"  {sku}: {quantidade} produto(s)")

    print("")
    print(f"Arquivo gerado: {ARQUIVO_SAIDA}")


if __name__ == "__main__":
    main()
