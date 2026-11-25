
from __future__ import annotations

import json
import unicodedata
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook


# ============================================================
# CONFIGURAÇÕES
# ============================================================

ROOT = Path(__file__).resolve().parents[1]

ARQUIVO_EXCEL = (
    ROOT / "data" / "unprocessed" / "Calendario Black Friday.xlsx"
)

ARQUIVO_SAIDA = (
    ROOT / "data" / "processed" / "black_friday_cerveja.json"
)

ABA = "CERVEJA"


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def normalizar_texto(valor) -> str:
    texto = str(valor or "").strip().lower()

    texto = unicodedata.normalize("NFKD", texto)

    return "".join(
        c for c in texto
        if not unicodedata.combining(c)
    )


def texto(valor) -> str:
    if valor is None:
        return ""

    return str(valor).strip()


def formatar_data(valor) -> str:
    if isinstance(valor, datetime):
        return valor.strftime("%d/%m/%Y")

    if isinstance(valor, date):
        return valor.strftime("%d/%m/%Y")

    return texto(valor)


def numero(valor):
    """
    Converte números do Excel preservando valores inteiros.
    Também aceita números escritos com vírgula decimal.
    """

    if valor is None or valor == "":
        return None

    if isinstance(valor, (int, float)):
        if isinstance(valor, float) and valor.is_integer():
            return int(valor)

        return valor

    valor = str(valor).strip().replace(",", ".")

    try:
        convertido = float(valor)

        if convertido.is_integer():
            return int(convertido)

        return convertido

    except ValueError:
        return valor


def valor_celula(ws, linha: int, coluna: int):
    return ws.cell(
        row=linha,
        column=coluna
    ).value


# ============================================================
# CONFIGURAÇÃO DAS MECÂNICAS
# ============================================================

def montar_configuracao(
    ws,
    linha: int,
    mecanica: str
) -> dict:

    mecanica_norm = normalizar_texto(mecanica)

    # --------------------------------------------------------
    # LEVE E PAGUE
    # I = LEVE
    # K = PAGUE
    # Q = MÁXIMO DE COMBOS POR CPF
    # --------------------------------------------------------

    if "leve" in mecanica_norm and "pague" in mecanica_norm:

        leve = numero(valor_celula(ws, linha, 9))
        pague = numero(valor_celula(ws, linha, 11))
        max_cpf = numero(valor_celula(ws, linha, 17))

        return {
            "tipo": "leve_pague",
            "leve": leve,
            "pague": pague,
            "max_por_cpf": max_cpf,
            "unidade_limite": "combos",
        }

    # --------------------------------------------------------
    # DESCONTO %
    # I = DESCONTO
    # K = MÍNIMO DE UNIDADES PARA ATIVAÇÃO
    # M = MÁXIMO DE UNIDADES POR CPF
    # --------------------------------------------------------

    if "desconto" in mecanica_norm:

        desconto = numero(valor_celula(ws, linha, 9))
        minimo = numero(valor_celula(ws, linha, 11))
        max_cpf = numero(valor_celula(ws, linha, 13))

        if isinstance(desconto, (int, float)):

            # A planilha apresenta 0.15 para representar 15%.
            if 0 < desconto <= 1:
                desconto *= 100

            desconto = round(desconto, 2)

        return {
            "tipo": "desconto",
            "desconto_percentual": desconto,
            "min_un_ativacao": minimo,
            "max_por_cpf": max_cpf,
            "unidade_limite": "unidades",
        }

    # --------------------------------------------------------
    # MECÂNICA NÃO MAPEADA
    # Não descartar silenciosamente a oferta.
    # --------------------------------------------------------

    return {
        "tipo": "nao_mapeada",
        "observacao": "Mecânica precisa de tratamento específico.",
    }


# ============================================================
# EXTRAÇÃO
# ============================================================

def extrair_ofertas(ws):

    ofertas = []
    problemas = []

    for linha in range(2, ws.max_row + 1):

        produto = texto(valor_celula(ws, linha, 6))

        if not produto:
            continue

        divisao = texto(valor_celula(ws, linha, 2))
        inicio = formatar_data(valor_celula(ws, linha, 3))
        fim = formatar_data(valor_celula(ws, linha, 4))
        mecanica = texto(valor_celula(ws, linha, 7))

        configuracao = montar_configuracao(
            ws,
            linha,
            mecanica
        )

        oferta = {
            "DIVISAO": divisao,
            "INICIO": inicio,
            "FIM": fim,
            "PRODUTO": produto,
            "MECANICA": mecanica,
            "CONFIGURACAO": configuracao,
        }

        ofertas.append(oferta)

        if configuracao["tipo"] == "nao_mapeada":
            problemas.append({
                "linha_excel": linha,
                "produto": produto,
                "mecanica": mecanica,
            })

    return ofertas, problemas


# ============================================================
# EXECUÇÃO
# ============================================================

def main():

    if not ARQUIVO_EXCEL.exists():
        raise FileNotFoundError(
            f"Arquivo não encontrado: {ARQUIVO_EXCEL}"
        )

    print("Abrindo arquivo Black Friday...")

    wb = load_workbook(
        ARQUIVO_EXCEL,
        data_only=True,
        read_only=False
    )

    try:

        if ABA not in wb.sheetnames:
            raise ValueError(
                f"A aba '{ABA}' não foi encontrada."
            )

        ws = wb[ABA]

        ofertas, problemas = extrair_ofertas(ws)

    finally:
        wb.close()

    ARQUIVO_SAIDA.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    ARQUIVO_SAIDA.write_text(
        json.dumps(
            ofertas,
            ensure_ascii=False,
            indent=2
        ),
        encoding="utf-8"
    )

    # --------------------------------------------------------
    # RESUMO
    # --------------------------------------------------------

    quantidade_leve_pague = sum(
        1 for o in ofertas
        if o["CONFIGURACAO"]["tipo"] == "leve_pague"
    )

    quantidade_desconto = sum(
        1 for o in ofertas
        if o["CONFIGURACAO"]["tipo"] == "desconto"
    )

    print("")
    print("=" * 60)
    print("EXTRAÇÃO BLACK FRIDAY CONCLUÍDA")
    print("=" * 60)

    print(f"Total de ofertas: {len(ofertas)}")
    print(f"Leve e Pague: {quantidade_leve_pague}")
    print(f"Desconto: {quantidade_desconto}")
    print(f"Não mapeadas: {len(problemas)}")

    print("")
    print(f"Arquivo gerado: {ARQUIVO_SAIDA}")

    if problemas:

        print("")
        print("MECÂNICAS QUE PRECISAM DE REVISÃO:")

        for problema in problemas:
            print(
                f"Linha {problema['linha_excel']}: "
                f"{problema['mecanica']}"
            )


if __name__ == "__main__":
    main()
