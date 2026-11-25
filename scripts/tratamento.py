from __future__ import annotations

import argparse
import json
import re
import unicodedata
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path, PureWindowsPath
from typing import Any

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data" / "unprocessed"
DEFAULT_OFFERS = DATA_DIR / "Calendario Mercapromo.xlsx"
DEFAULT_CONFIG_BOOK = DATA_DIR / "Calendario Mercapromo.xlsx"
DEFAULT_MASTER = ROOT / "cadastro" / "de_para_mercapromo.xlsx"
DEFAULT_OUTPUT = ROOT / "data" / "processed" / "calendario_mercapromo.csv"
DEFAULT_REPORT = ROOT / "data" / "processed" / "relatorio_tratamento.txt"
DEFAULT_META = ROOT / "data" / "processed" / "calendario_mercapromo_meta.json"

MONTHS = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
    7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
}
MONTH_NUMBERS = {re.sub(r"[^a-z]", "", unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode("ascii").casefold()): number for number, name in MONTHS.items()}
CATEGORIES = ("CERVEJA", "NAB", "MATCH", "PERSONALIZADO")
CATEGORY_ALIASES = {
    "cerveja": "CERVEJA", "cerv": "CERVEJA", "nab": "NAB", "match": "MATCH",
    "personalizado": "PERSONALIZADO", "mercapromo": "PERSONALIZADO",
}
MASTER_TYPE = {"CERVEJA": "CERV", "NAB": "NAB", "MATCH": "MATCH", "PERSONALIZADO": "MERCAPROMO"}
OUTPUT_COLUMNS = [
    "categoria", "mes", "ano", "inicio", "fim", "quinzena", "produto_origem", "mecanica",
    "tipo_mecanica", "quantidade_mecanica", "rotulo_ativacao", "valor_ativacao", "max_combo_cpf", "mecanica_visual",
    "divisao", "sinalizador_variacao", "varejista", "abrangencia", "tipo_de_para", "promax", "eg", "nome",
    "nome_abreviado", "nome_imagem", "caminho_imagem", "nome_imagem_secundaria",
    "caminho_imagem_secundaria", "oferta_manual", "mapeado",
]

# Expansões dos prefixos mais comuns usados no SKU visual do De/Para.
# Elas permitem relacionar o nome comercial longo da oferta (ex. "Antarctica Pilsen LT 269ml")
# ao SKU abreviado que nomeia a imagem (ex. "AP 269"). O casamento só é aceito quando marca,
# tamanho e atributos de variante são compatíveis; em caso de ambiguidade, o script não escolhe.
SKU_PREFIX_EXPANSIONS: dict[str, tuple[str, ...]] = {
    "bc0": ("brahma", "chopp", "zero"), "bcz": ("brahma", "chopp", "zero"),
    "budz": ("budweiser", "zero"), "skz": ("skol", "zero"),
    "ap": ("antarctica", "pilsen"), "bc": ("brahma", "chopp"),
    "bud": ("budweiser",), "sk": ("skol",), "cor": ("corona",),
    "ori": ("original",), "boh": ("bohemia",), "spat": ("spaten",),
    "mich": ("michelob",), "mic": ("michelob",), "beats": ("beats",), "ff": ("flying", "fish"),
    "gca": ("guarana", "antarctica"), "pep": ("pepsi",), "gat": ("gatorade",),
    "h2o": ("h2oh",), "bare": ("bare",), "fuse": ("fusion",),
    "pc": ("pepsi",), "gatorade": ("gatorade",),
}
KNOWN_PACK_SIZES = (269, 275, 330, 350, 355, 473, 500, 550, 600, 1000, 1500, 2000)

# Mapeamento por posição física (zero-based), não pelo nome dos cabeçalhos.
# Nas abas CERVEJA/NAB/MATCH, H:K podem estar sob um cabeçalho mesclado.
SHEET_SPECS: dict[str, dict[str, Any]] = {
    "CERVEJA": {
        "division": 1, "start": 2, "end": 3, "product": 5, "mechanic": 6,
        "kind": 7, "quantity": 8, "activation_label": 9, "activation_value": 10,
        "retailer": None, "max_combo": 16, "first_data_row": 2, "usecols": "A:Q", "fallback_usecols": "A:K",
    },
    "NAB": {
        "division": 1, "start": 2, "end": 3, "product": 5, "mechanic": 6,
        "kind": 7, "quantity": 8, "activation_label": 9, "activation_value": 10,
        "retailer": None, "max_combo": 16, "first_data_row": 2, "usecols": "A:Q", "fallback_usecols": "A:K",
    },
    "MATCH": {
        "division": 1, "start": 2, "end": 3, "product": 5, "mechanic": 6,
        "kind": 7, "quantity": 8, "activation_label": 9, "activation_value": 10,
        "retailer": None, "max_combo": 16, "first_data_row": 2, "usecols": "A:Q", "fallback_usecols": "A:K",
    },
    "PERSONALIZADO": {
        "division": 0, "retailer": 2, "start": 3, "end": 4, "product": 6, "mechanic": 7,
        "kind": 8, "quantity": 9, "activation_label": 10, "activation_value": 11,
        "max_combo": 13, "first_data_row": 3, "usecols": "A:N", "fallback_usecols": "A:L",
    },
}



def clean(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, (int,)):
        return str(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    return "" if text.casefold() in {"nan", "none", "null", "nat", "-"} else text


def norm(value: Any) -> str:
    text = unicodedata.normalize("NFKD", clean(value))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]+", "", text.casefold())


UF_CODES = ("AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS", "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN", "RS", "RO", "RR", "SC", "SP", "SE", "TO")
UF_PATTERN = "(?:" + "|".join(UF_CODES) + ")"
RETAILER_SCOPE_RE = re.compile(
    rf"(?i)\s+(?P<scope>{UF_PATTERN}(?:\s*(?:e|/|,|&)\s*{UF_PATTERN})*)\s*$"
)
# Certain suffixes describe the scope/channel of a retailer variant, not a separate
# retail network. Keep the network filter consolidated (e.g. Fort rather than Fort Atacado).
RETAILER_CHANNEL_RE = re.compile(r"(?i)\s+(?P<scope>ATACADO|VAREJO|DIRETA|VIZINHANCA|ON\s+TRADE)\s*$")


def split_retailer_scope(value: Any) -> tuple[str, str]:
    """Split suffixes such as 'Comper DF e GO' or 'Fort Atacado' into network and badge."""
    retailer = clean(value)
    if not retailer:
        return "", ""
    # Fort is an atacado-only network, so the channel is implicit: consolidate
    # "Fort Atacado" into the single filter "Fort" without displaying *Atacado.
    if norm(retailer) == "fortatacado":
        return "Fort", ""
    match = RETAILER_SCOPE_RE.search(retailer)
    if match:
        network = retailer[:match.start()].strip(" -_,/;\t")
        scope = match.group("scope").strip()
        # Present mixed separators consistently without changing state codes.
        scope = re.sub(r"\s*[/,&]\s*", " e ", scope)
        scope = re.sub(r"\s+e\s+", " e ", scope, flags=re.IGNORECASE).upper()
        scope = re.sub(r"\bE\b", "e", scope)
        return (network or retailer), scope
    channel_match = RETAILER_CHANNEL_RE.search(retailer)
    if channel_match:
        network = retailer[:channel_match.start()].strip(" -_,/;\t")
        scope = re.sub(r"\s+", " ", channel_match.group("scope").strip()).title()
        return (network or retailer), scope
    return retailer, ""


def canonical_category(value: Any) -> str:
    return CATEGORY_ALIASES.get(norm(value), clean(value).upper())


def canonical_month(value: Any) -> str:
    text = clean(value)
    if not text:
        return ""
    numeric = pd.to_numeric(pd.Series([text]), errors="coerce").iloc[0]
    if pd.notna(numeric) and int(numeric) in MONTHS:
        return MONTHS[int(numeric)]
    return MONTHS.get(MONTH_NUMBERS.get(norm(text), 0), text[:1].upper() + text[1:])


def timestamp(value: Any) -> pd.Timestamp | None:
    if isinstance(value, (datetime, date, pd.Timestamp)):
        return pd.Timestamp(value).normalize()
    text = clean(value)
    if not text:
        return None
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}(?:[ T].*)?", text):
        parsed = pd.to_datetime(text, errors="coerce")
    else:
        parsed = pd.to_datetime(text, dayfirst=True, errors="coerce")
    return None if pd.isna(parsed) else pd.Timestamp(parsed).normalize()


def parse_date(value: Any) -> str:
    parsed = timestamp(value)
    return parsed.strftime("%Y-%m-%d") if parsed is not None else ""


def month_bounds(month: str, year: int) -> tuple[pd.Timestamp, pd.Timestamp]:
    month_no = next((n for n, name in MONTHS.items() if norm(name) == norm(month)), None)
    if month_no is None:
        raise ValueError(f"Mês vigente inválido: {month!r}")
    first = pd.Timestamp(year=year, month=month_no, day=1)
    return first, first + pd.offsets.MonthEnd(0)


def overlaps_month(start: Any, end: Any, first: pd.Timestamp, last: pd.Timestamp) -> bool:
    a, b = timestamp(start), timestamp(end)
    return a is not None and b is not None and min(a, b) <= last and max(a, b) >= first


def category_quarter(start: str, end: str) -> str:
    a, b = timestamp(start), timestamp(end)
    if a is None or b is None:
        return "PERÍODO"
    # No calendário real as janelas podem ser 01–08 e 09–29 (não necessariamente
    # as quinzenas 01–15 e 16–fim). Usamos a data inicial para nomear a janela.
    return "1ª QUINZENA" if a.day <= 8 else "2ª QUINZENA"


def format_numeric(value: Any) -> str:
    text = clean(value)
    if not text:
        return ""
    # Normalize numbers such as 6.0 without destroying text labels.
    numeric = pd.to_numeric(pd.Series([text.replace(",", ".")]), errors="coerce").iloc[0]
    if pd.notna(numeric):
        return str(int(numeric)) if float(numeric).is_integer() else str(numeric).replace(".", ",")
    return text


def format_discount(value: Any) -> str:
    text = clean(value)
    if not text:
        return ""
    if "%" in text:
        return text
    numeric = pd.to_numeric(pd.Series([text.replace(",", ".")]), errors="coerce").iloc[0]
    if pd.isna(numeric):
        return text
    number = float(numeric)
    # Excel percentages can arrive as 0.2 or as 20 depending on cell formatting.
    if 0 < number < 1:
        number *= 100
    display = str(int(number)) if number.is_integer() else str(number).replace(".", ",")
    return f"{display}%"


def display_mechanic(kind: Any, quantity: Any, activation_label: Any, activation_value: Any, description: Any = "") -> str:
    kind_text = clean(kind).rstrip(":").strip()
    qty = clean(quantity)
    label = clean(activation_label).rstrip(":").strip()
    value = clean(activation_value)
    key = norm(kind_text)
    if key.startswith("leve"):
        amount = format_numeric(qty)
        pay = format_numeric(value)
        first = f"Leve {amount}" if amount else "Leve"
        if pay:
            second = f"{label or 'Pague'} {pay}"
            return first + " / " + second
        return first
    if key.startswith("desconto"):
        percent = format_discount(qty)
        first = f"Desconto {percent}" if percent else "Desconto"
        threshold = format_numeric(value)
        if threshold:
            label_key = norm(label)
            if "minim" in label_key or "unidade" in label_key or "ativacao" in label_key or "item" in label_key:
                if threshold == "1":
                    second = "Mínimo de 1 item"
                else:
                    second = f"Mínimo de {threshold} itens"
            else:
                second = f"{label or 'A partir de'} {threshold} itens"
            return first + " / " + second
        return first
    # Special promotions can put their complete copy in MECÂNICA column G/H.
    return clean(description) or "Mecânica não especificada"


def read_offer_columns(workbook: pd.ExcelFile, sheet: str, spec: dict[str, Any]) -> pd.DataFrame:
    """Read the requested physical Excel range; maximum-combo columns are optional in small templates."""
    try:
        return pd.read_excel(
            workbook, sheet_name=sheet, header=None, dtype=object,
            usecols=spec["usecols"],
        )
    except pd.errors.ParserError:
        fallback = spec.get("fallback_usecols")
        if not fallback or fallback == spec["usecols"]:
            raise
        return pd.read_excel(
            workbook, sheet_name=sheet, header=None, dtype=object,
            usecols=fallback,
        )


def infer_promotion_month(offers_path: Path) -> tuple[str, int, Counter[tuple[int, int]]]:
    """Infer the campaign month from the four active offer tabs, using physical columns."""
    counts: Counter[tuple[int, int]] = Counter()
    workbook = pd.ExcelFile(offers_path)
    sheet_map = {norm(name): name for name in workbook.sheet_names}
    for category, spec in SHEET_SPECS.items():
        sheet = sheet_map.get(norm(category))
        if not sheet:
            continue
        raw = read_offer_columns(workbook, sheet, spec)
        for excel_row_no, (_, row) in enumerate(raw.iterrows(), start=1):
            if excel_row_no < spec["first_data_row"]:
                continue
            product_pos, start_pos, end_pos = spec["product"], spec["start"], spec["end"]
            if max(product_pos, start_pos, end_pos) >= len(row):
                continue
            product = clean(row.iloc[product_pos])
            start_date, end_date = timestamp(row.iloc[start_pos]), timestamp(row.iloc[end_pos])
            if product and norm(product) not in {"produto", "produtos", "sku"} and start_date is not None and end_date is not None:
                counts[(start_date.year, start_date.month)] += 1

    now = datetime.now()
    current_key = (now.year, now.month)
    if current_key in counts:
        chosen = current_key
    else:
        future_keys = [key for key in counts if key >= current_key]
        if future_keys:
            # If there is no data for the current month, choose the most represented upcoming campaign month.
            chosen = sorted(future_keys, key=lambda key: (-counts[key], key[0], key[1]))[0]
        elif counts:
            chosen = max(counts, key=lambda key: (counts[key], key[0], key[1]))
        else:
            chosen = current_key
    return MONTHS[chosen[1]], chosen[0], counts


def read_config(config_book: Path | None, offers_path: Path | None = None) -> tuple[str, int, dict[str, str], list[str]]:
    now = datetime.now()
    month, year = MONTHS[now.month], now.year
    counts: Counter[tuple[int, int]] = Counter()
    warnings: list[str] = []
    config: dict[str, str] = {}
    if offers_path and offers_path.is_file():
        month, year, counts = infer_promotion_month(offers_path)

    found_config = False
    if config_book and config_book.is_file():
        try:
            wb = pd.ExcelFile(config_book)
            sheets = {norm(name): name for name in wb.sheet_names}
            if "config" in sheets:
                df = pd.read_excel(wb, sheet_name=sheets["config"], header=0, dtype=object, usecols="A:B")
                if len(df.columns) >= 2:
                    key_col, val_col = df.columns[0], df.columns[1]
                    for _, row in df.iterrows():
                        key, value = clean(row.get(key_col, "")), clean(row.get(val_col, ""))
                        if key:
                            config[norm(key)] = value
                    found_config = True
                    if config.get("mesatual"):
                        month = canonical_month(config["mesatual"])
                    if config.get("anoatual"):
                        value = pd.to_numeric(pd.Series([config["anoatual"]]), errors="coerce").iloc[0]
                        if pd.notna(value):
                            year = int(value)
            else:
                warnings.append(f"Aba CONFIG não existe em {config_book.name}; mês inferido pelas datas das ofertas.")
        except Exception as exc:
            warnings.append(f"Não consegui ler CONFIG ({exc}); mês inferido pelas datas das ofertas.")
    else:
        warnings.append("CONFIG não encontrado; mês inferido pelas datas das ofertas.")

    if not found_config:
        if counts:
            summary = ", ".join(f"{MONTHS[m]} {y}: {n} linha(s)" for (y, m), n in sorted(counts.items()))
            warnings.append(f"Distribuição de datas por mês: {summary}. Selecionado: {month}/{year}.")
        else:
            warnings.append(f"Não foi possível inferir o mês; usando o mês do sistema ({month}/{year}).")
    return month, year, config, warnings

def find_offers_file(explicit: Path | None) -> Path:
    """Find the workbook by the required sheet names; columns are always read by physical position."""
    required = {norm(x) for x in CATEGORIES}

    def sheet_names(path: Path) -> list[str]:
        try:
            return list(pd.ExcelFile(path).sheet_names)
        except Exception:
            return []

    def valid(path: Path) -> bool:
        names = sheet_names(path)
        normalized = {norm(sheet) for sheet in names}
        return required.issubset(normalized)

    if explicit is not None:
        candidate = explicit if explicit.is_absolute() else (ROOT / explicit)
        if not candidate.is_file():
            raise FileNotFoundError(f"Planilha informada não encontrada: {candidate}")
        names = sheet_names(candidate)
        if not valid(candidate):
            raise ValueError(
                f"A planilha {candidate.name} precisa conter as abas CERVEJA, NAB, MATCH e PERSONALIZADO. "
                f"Abas encontradas: {names}"
            )
        return candidate

    if DEFAULT_OFFERS.is_file() and valid(DEFAULT_OFFERS):
        return DEFAULT_OFFERS

    candidates: list[Path] = []
    if DATA_DIR.is_dir():
        candidates = sorted(
            p for p in DATA_DIR.iterdir()
            if p.is_file() and p.suffix.lower() in {".xlsx", ".xlsm", ".xlsb"}
        )
    for path in candidates:
        if valid(path):
            return path

    diagnostics = []
    for path in candidates:
        names = sheet_names(path)
        diagnostics.append(f"- {path.name}: {', '.join(names) if names else 'não foi possível ler as abas'}")
    detail = "\n".join(diagnostics) if diagnostics else "Nenhum Excel encontrado em data/unprocessed/."
    raise FileNotFoundError(
        "Não encontrei uma planilha com as abas CERVEJA, NAB, MATCH e PERSONALIZADO em data/unprocessed/. "
        f"A planilha esperada é {DEFAULT_OFFERS.name}.\nAbas dos arquivos encontrados:\n{detail}"
    )


def read_master(master_file: Path | None, offers_path: Path) -> tuple[pd.DataFrame, str, list[str]]:
    warnings: list[str] = []
    if master_file and master_file.is_file():
        return pd.read_excel(master_file, dtype=object), str(master_file), warnings
    if DEFAULT_MASTER.is_file():
        return pd.read_excel(DEFAULT_MASTER, dtype=object), str(DEFAULT_MASTER), warnings
    try:
        wb = pd.ExcelFile(offers_path)
        sheet = next((s for s in wb.sheet_names if norm(s) == "deparaimg"), None)
        if sheet:
            warnings.append(f"De/Para externo ausente; usando a aba {sheet} de {offers_path.name}.")
            return pd.read_excel(wb, sheet_name=sheet, dtype=object), f"{offers_path.name} / {sheet}", warnings
    except Exception:
        pass
    warnings.append("De/Para não encontrado. Os cards serão mantidos sem Promax/imagem até o cadastro ser incluído.")
    return pd.DataFrame(), "não encontrado", warnings


def build_lookup(master: pd.DataFrame) -> tuple[dict[str, list[dict[str, str]]], list[str]]:
    """Build aliases from the De/Para. `Nome` is the canonical product description,
    `SKU` is the abbreviated label, and `Nome IMG`/`Caminho` identify the image asset.
    """
    by_alias: dict[str, list[dict[str, str]]] = defaultdict(list)
    warnings: list[str] = []
    if master.empty:
        return {}, warnings
    cols = {norm(c): c for c in master.columns}

    def column(canonical: str, pos: int) -> Any:
        return cols.get(norm(canonical), master.columns[pos] if len(master.columns) > pos else None)

    c_type, c_promax, c_eg, c_sku, c_img, c_path = [
        column(name, i) for i, name in enumerate(["Tipo", "Promax", "EG", "SKU", "Nome IMG", "Caminho"])
    ]
    # `Nome` was added to the De/Para after the first version. Do not use a positional
    # fallback here, because old six-column files may not contain it.
    c_name = cols.get(norm("Nome"))

    for _, row in master.iterrows():
        row_type = clean(row.get(c_type, "")) if c_type is not None else ""
        promax = clean(row.get(c_promax, "")) if c_promax is not None else ""
        eg = clean(row.get(c_eg, "")) if c_eg is not None else ""
        sku = clean(row.get(c_sku, "")) if c_sku is not None else ""
        img = clean(row.get(c_img, "")) if c_img is not None else ""
        path = clean(row.get(c_path, "")) if c_path is not None else ""
        full_name = clean(row.get(c_name, "")) if c_name is not None else ""
        if norm(sku) == "sku" or norm(full_name) == "nome" or not any((sku, eg, img, path, promax, full_name)):
            continue

        # SKU is the visible short name; Nome is for matching the promotion's long name.
        # Nome IMG and Caminho are used to find the physical image file.
        path_stem = PureWindowsPath(path.replace("/", "\\")).stem if path else ""
        img_stem = PureWindowsPath(img.replace("/", "\\")).stem if img else ""
        info = {
            "tipo_de_para": row_type,
            "promax": promax,
            "eg": eg,
            "nome": full_name or sku or eg or promax,
            "nome_abreviado": sku or eg or promax,
            "nome_imagem": img_stem or img,
            "caminho_imagem": path,
        }
        aliases = (full_name, sku, eg, promax, path, path_stem, img, img_stem)
        for alias in aliases:
            key = norm(alias)
            if key:
                by_alias[key].append(info)

    for alias, candidates in by_alias.items():
        promaxes = {c["promax"] for c in candidates if c.get("promax") and norm(c["promax"]) != ""}
        if len(promaxes) > 1:
            warnings.append(
                f"Alias {alias!r} associado a Promax diferentes ({', '.join(sorted(promaxes))}); "
                "a associação exata ambígua não será escolhida automaticamente."
            )
    return dict(by_alias), warnings


def extract_volume_ml(value: Any) -> int | None:
    """Read consumer package size from product descriptions (269ml, 1,5L, 2 litros)."""
    raw = clean(value).casefold().replace("\u00a0", " ")
    if not raw:
        return None
    match = re.search(r"(?<!\d)(\d{2,4})\s*ml\b", raw)
    if match:
        return int(match.group(1))
    match = re.search(r"(?<!\d)(\d+(?:[.,]\d+)?)\s*(?:litros?|l)\b", raw)
    if match:
        try:
            liters = float(match.group(1).replace(",", "."))
            if 0 < liters <= 5:
                return int(round(liters * 1000))
        except ValueError:
            pass
    normalized = norm(raw)
    for size in KNOWN_PACK_SIZES:
        if re.search(rf"(?<!\d){size}(?!\d)", normalized):
            return size
    return None


def sku_code_volume_ml(sku: str) -> int | None:
    """Interpret volume codes like P1, P1,5 and P2 used by some De/Para SKUs."""
    normalized = norm(sku)
    # Explicit ml-size codes take priority over P1/P2 package codes.
    for size in sorted(KNOWN_PACK_SIZES, reverse=True):
        if re.search(rf"(?<!\d){size}(?!\d)", normalized):
            return size
    match = re.search(r"p(15|1|2|3|4)(?:mp)?$", normalized)
    if match:
        return {"1": 1000, "15": 1500, "2": 2000, "3": 3000, "4": 4000}[match.group(1)]
    return None


BEATS_VARIANTS: dict[str, tuple[str, ...]] = {
    "caip": ("caipirinha",), "caipirinha": ("caipirinha",),
    "trop": ("tropical",), "tropical": ("tropical",),
    "moj": ("mojito",), "mojito": ("mojito",),
    "mos": ("moscowmule",), "moscowmule": ("moscowmule",),
    "sen": ("senses",), "senses": ("senses",),
    "red": ("redmix", "red"), "redmix": ("redmix",),
    "green": ("greenmix", "green"), "greenmix": ("greenmix",),
    "gt": ("gt",),
}


def resolve_product(product: str, category: str, lookup: dict[str, list[dict[str, str]]]) -> dict[str, str] | None:
    """Link the promotion's long product name to the De/Para SKU, image name and path."""
    product_key = norm(product)
    candidates = lookup.get(product_key, [])
    target = MASTER_TYPE.get(category, category)

    def signature(item: dict[str, str]) -> tuple[str, ...]:
        return tuple(clean(item.get(k, "")) for k in (
            "tipo_de_para", "promax", "eg", "nome", "nome_abreviado", "nome_imagem", "caminho_imagem"
        ))

    def unique_or_none(values: list[dict[str, str]]) -> dict[str, str] | None:
        unique = {signature(item): item for item in values}
        return next(iter(unique.values())) if len(unique) == 1 else None

    def typed_pool(values: list[dict[str, str]]) -> list[dict[str, str]]:
        typed = [item for item in values if norm(item.get("tipo_de_para", "")) in {norm(target), norm(category)}]
        return typed or values

    if candidates:
        exact = unique_or_none(typed_pool(candidates))
        if exact:
            return exact
        # Don't replace a genuine exact alias collision with a guess.
        return None

    # The De/Para `Nome` column contains the canonical, often longer product name
    # (e.g. "ANTARCTICA PILSEN LT 269ML SH TRANSP C/15 NPAL"), whereas promotion
    # tabs may say only "Antarctica Pilsen LT 269ml". Match by containment first.
    # This makes the relationship explicit instead of relying on SKU-prefix guesses.
    full_name_matches: list[tuple[int, dict[str, str]]] = []
    full_name_pool: dict[tuple[str, ...], dict[str, str]] = {}
    for aliases in lookup.values():
        for item in aliases:
            full_name_pool[signature(item)] = item
    product_volume = extract_volume_ml(product)
    product_raw = clean(product).casefold()
    product_form = "ln" if re.search(r"\bln\b", product_raw) else "lt" if re.search(r"\blt\b", product_raw) else ""
    product_has_zero = any(token in product_key for token in ("zero", "cero", "semalcool", "0alcohol"))
    product_is_multipack = any(token in product_key for token in ("multipack", "multipac", "pack"))
    product_variant_terms = {term for terms in BEATS_VARIANTS.values() for term in terms if norm(term) in product_key}

    for item in full_name_pool.values():
        item_type = norm(item.get("tipo_de_para", ""))
        if category != "PERSONALIZADO" and item_type not in {norm(target), norm(category)}:
            continue
        full_name = clean(item.get("nome", ""))
        full_key = norm(full_name)
        if not full_key:
            continue
        if product_volume is not None:
            candidate_volume = extract_volume_ml(full_name) or sku_code_volume_ml(item.get("nome_abreviado", ""))
            if candidate_volume is not None and candidate_volume != product_volume:
                continue
        full_raw = full_name.casefold()
        full_form = "ln" if re.search(r"\bln\b", full_raw) else "lt" if re.search(r"\blt\b", full_raw) else ""
        if product_form and full_form and product_form != full_form:
            continue
        full_has_zero = any(token in full_key for token in ("zero", "cero", "semalcool", "0alcohol"))
        if product_has_zero != full_has_zero and (product_has_zero or full_has_zero):
            continue
        full_is_multipack = any(token in full_key for token in ("multipack", "multipac", "pack")) or norm(item.get("nome_abreviado", "")).endswith("mp")
        if product_is_multipack != full_is_multipack and (product_is_multipack or full_is_multipack):
            continue
        sku_key = norm(item.get("nome_abreviado", ""))
        if product_variant_terms and sku_key.startswith("beats"):
            sku_variants = [(marker, terms) for marker, terms in BEATS_VARIANTS.items() if marker in sku_key]
            if sku_variants and not any(norm(term) in product_key for term in sku_variants[0][1]):
                continue
            if not sku_variants and not any(norm(term) in full_key for term in product_variant_terms):
                continue

        if product_key in full_key:
            # Prefer the most specific name with the smallest extra suffix.
            score = 20000 - (len(full_key) - len(product_key))
        elif full_key in product_key:
            # Promotion text may be longer than the De/Para name; prefer the longest match.
            score = 19000 + len(full_key)
        else:
            continue
        full_name_matches.append((score, item))

    if full_name_matches:
        best_score = max(score for score, _ in full_name_matches)
        best = [item for score, item in full_name_matches if score == best_score]
        named = unique_or_none(best)
        if named:
            return named
        # If more than one De/Para row is equally specific, do not guess.
        return None

    # Deduplicate rows because each De/Para record is indexed under multiple aliases.
    pool: dict[tuple[str, ...], dict[str, str]] = {}
    for aliases in lookup.values():
        for item in aliases:
            pool[signature(item)] = item

    product_volume = extract_volume_ml(product)
    product_norm = product_key
    product_form = "ln" if re.search(r"\bln\b", clean(product).casefold()) else "lt" if re.search(r"\blt\b", clean(product).casefold()) else ""
    product_has_zero = any(token in product_norm for token in ("zero", "cero", "semalcool", "0alcohol"))
    product_is_multipack = any(token in product_norm for token in ("multipack", "multipac", "pack"))

    matches: list[tuple[int, dict[str, str]]] = []
    for item in pool.values():
        if category != "PERSONALIZADO" and norm(item.get("tipo_de_para", "")) not in {norm(target), norm(category)}:
            continue
        sku = clean(item.get("nome_abreviado", ""))
        sku_norm = norm(sku)
        if not sku_norm:
            continue

        expansion = next((terms for prefix, terms in sorted(SKU_PREFIX_EXPANSIONS.items(), key=lambda entry: -len(entry[0]))
                          if sku_norm.startswith(prefix)), None)
        if not expansion or not all(norm(term) in product_norm for term in expansion):
            continue

        score = 10 * len(expansion)
        sku_volume = sku_code_volume_ml(sku)
        if sku_volume is not None and product_volume is not None:
            if sku_volume != product_volume:
                continue
            score += 35
        elif sku_volume is not None and product_volume is None:
            # No volume to compare: only use this SKU when it is the only viable choice later.
            score -= 8
        elif product_volume is not None:
            # A generic image SKU may still be valid, but loses to an exact-size SKU.
            score -= 2

        sku_form = "ln" if re.search(r"ln", sku_norm) else "lt" if re.search(r"lt", sku_norm) else ""
        if sku_form and product_form and sku_form != product_form:
            continue
        if sku_form and product_form and sku_form == product_form:
            score += 8

        sku_has_zero = sku_norm.startswith(("bc0", "bcz", "budz", "skz", "corcz", "corcero")) or "zero" in sku_norm
        if sku_has_zero and not product_has_zero:
            continue
        if product_has_zero and not sku_has_zero and sku_norm.startswith(("bc", "bud", "sk", "cor")):
            # Prevent a named Zero variant being linked to the regular SKU.
            continue
        if sku_has_zero and product_has_zero:
            score += 8

        sku_is_multipack = sku_norm.endswith("mp") or "multipack" in sku_norm
        if sku_is_multipack and not product_is_multipack:
            continue
        if sku_is_multipack and product_is_multipack:
            score += 12
        elif product_is_multipack:
            score -= 5

        # For Beats, enforce any specific flavor encoded in the SKU. A generic
        # "Beats Sabores" promotion should map to a generic BEATS SKU, not Caipirinha/Mojito/etc.
        if sku_norm.startswith("beats"):
            sku_variants = [(marker, terms) for marker, terms in BEATS_VARIANTS.items() if marker in sku_norm]
            product_variant_terms = {term for terms in BEATS_VARIANTS.values() for term in terms if norm(term) in product_norm}
            if sku_variants:
                required = sku_variants[0][1]
                if not any(norm(term) in product_norm for term in required):
                    continue
                score += 25
            elif product_variant_terms:
                # Generic SKU can be a fallback for a variant, but only if there is no exact variant SKU.
                score -= 4

        if "ln" in sku_norm and "ln" in product_norm:
            score += 2
        if "lt" in sku_norm and "lt" in product_norm:
            score += 2
        matches.append((score, item))

    if not matches:
        return None
    best_score = max(score for score, _ in matches)
    best = [item for score, item in matches if score == best_score]
    return unique_or_none(best)

def mechanic_from_row(row: pd.Series, spec: dict[str, int]) -> dict[str, str]:
    def value(key: str) -> Any:
        idx = spec.get(key)
        return row.iloc[idx] if idx is not None and idx < len(row) else ""
    kind = clean(value("kind"))
    quantity = clean(value("quantity"))
    label = clean(value("activation_label"))
    activation_value = clean(value("activation_value"))
    desc = clean(value("mechanic"))
    return {
        "mecanica": desc,
        "tipo_mecanica": kind,
        "quantidade_mecanica": quantity,
        "rotulo_ativacao": label,
        "valor_ativacao": activation_value,
        "max_combo_cpf": clean(value("max_combo")),
        "mecanica_visual": display_mechanic(kind, quantity, label, activation_value, desc),
    }


def transform(offers_file: Path | None, master_file: Path | None, config_file: Path | None,
              output_file: Path, report_file: Path, meta_file: Path) -> pd.DataFrame:
    offers_path = find_offers_file(offers_file)
    config_path = config_file if config_file is not None else (DEFAULT_CONFIG_BOOK if DEFAULT_CONFIG_BOOK.is_file() else None)
    month, year, config, warnings = read_config(config_path, offers_path)
    first, last = month_bounds(month, year)
    master_df, master_description, master_warnings = read_master(master_file, offers_path)
    lookup, lookup_warnings = build_lookup(master_df)
    warnings.extend(master_warnings)
    warnings.extend(lookup_warnings)

    workbook = pd.ExcelFile(offers_path)
    sheet_map = {norm(s): s for s in workbook.sheet_names}
    specs = SHEET_SPECS

    raw_rows: list[dict[str, Any]] = []
    excluded_month = excluded_invalid = empty_rows = 0
    for category, spec in specs.items():
        sheet = sheet_map.get(norm(category))
        if not sheet:
            warnings.append(f"Aba obrigatória {category} não encontrada em {offers_path.name}.")
            continue
        # Limita a leitura às colunas físicas necessárias: a planilha contém colunas auxiliares
        # e formatação até XFD, que não fazem parte do calendário visual.
        raw = read_offer_columns(workbook, sheet, spec)
        for excel_row_no, (_, row) in enumerate(raw.iterrows(), start=1):
            if excel_row_no < spec["first_data_row"]:
                continue
            def val(key: str) -> Any:
                pos = spec.get(key)
                return row.iloc[pos] if pos is not None and pos < len(row) else ""
            product = clean(val("product"))
            if not product or norm(product) in {"produto", "produtos", "sku", "nomeproduto"}:
                empty_rows += 1
                continue
            start_ts, end_ts = timestamp(val("start")), timestamp(val("end"))
            if start_ts is None or end_ts is None:
                excluded_invalid += 1
                continue
            if not overlaps_month(start_ts, end_ts, first, last):
                excluded_month += 1
                continue
            start_text, end_text = start_ts.strftime("%Y-%m-%d"), end_ts.strftime("%Y-%m-%d")
            # Product map is based on exact product name against De/Para SKU, EG or Promax.
            mapping = resolve_product(product, category, lookup)
            division = clean(val("division"))
            retailer_raw = clean(val("retailer"))
            retailer, retailer_scope = split_retailer_scope(retailer_raw) if category == "PERSONALIZADO" else (retailer_raw, "")
            mechanic = mechanic_from_row(row, spec)
            quarter = category_quarter(start_text, end_text)
            info = mapping or {}

            # Regra de negócio: todo produto NAB tem limite de 1 por CPF em todas
            # as quinzenas, inclusive quando aparece dentro de PERSONALIZADO.
            de_para_type = norm(info.get("tipo_de_para", ""))
            product_identity = norm(" ".join([
                product,
                clean(info.get("nome", "")),
                clean(info.get("nome_abreviado", "")),
            ]))
            # Applied to every row in the NAB tab and to personalized NAB products.
            # If a personalized product is missing from the De/Para, known NAB brands
            # still inherit the fixed limit rather than keeping an Excel default.
            nab_product_markers = (
                "guarana", "pepsi", "gatorade", "fusion", "h2oh", "bare p", "barep",
                "agua", "aguamineral", "tonica", "sodawater",
            )
            is_nab_offer = category == "NAB" or (
                category == "PERSONALIZADO"
                and (
                    de_para_type in {"nab", "naoalcoolica", "bebidanaoalcoolica"}
                    or any(norm(marker) in product_identity for marker in nab_product_markers)
                )
            )
            if is_nab_offer:
                mechanic["max_combo_cpf"] = "1"

            # Regra de negócio: todas as ofertas regulares MATCH têm limite de
            # 4 combos por CPF nas duas quinzenas, inclusive no PERSONALIZADO quando
            # o De/Para classifica o produto como MATCH. A ação especial BEATS - FF
            # é tratada abaixo e tem exceção própria de 3 combos por CPF.
            is_match_offer = category == "MATCH" or (
                category == "PERSONALIZADO"
                and de_para_type in {"match", "bebidamista"}
            )
            if is_match_offer:
                mechanic["max_combo_cpf"] = "4"
            raw_rows.append({
                "categoria": category, "mes": month, "ano": str(year), "inicio": start_text, "fim": end_text,
                "quinzena": quarter, "produto_origem": product,
                **mechanic, "divisao": division, "sinalizador_variacao": "", "varejista": retailer, "abrangencia": retailer_scope,
                "tipo_de_para": info.get("tipo_de_para", ""), "promax": info.get("promax", ""), "eg": info.get("eg", ""),
                "nome": info.get("nome", "") or product, "nome_abreviado": info.get("nome_abreviado", "") or product,
                "nome_imagem": info.get("nome_imagem", ""), "caminho_imagem": info.get("caminho_imagem", ""),
                "nome_imagem_secundaria": "", "caminho_imagem_secundaria": "", "oferta_manual": "não",
                "mapeado": "sim" if mapping else "não",
            })

    # Oferta manual de cross-sell pedida pelo negócio. Não existe como produto na base,
    # então a criamos em MATCH na PRIMEIRA quinzena do calendário mensal.
    match_windows = sorted({
        (clean(item.get("inicio", "")), clean(item.get("fim", "")))
        for item in raw_rows if item.get("categoria") == "MATCH" and item.get("inicio") and item.get("fim")
    })
    first_windows = [window for window in match_windows
                     if timestamp(window[0]) is not None and category_quarter(window[0], window[1]) == "1ª QUINZENA"]
    chosen_window = first_windows[0] if first_windows else (match_windows[0] if match_windows else ("", ""))
    existing_manual = any(
        item["categoria"] == "MATCH" and norm(item.get("produto_origem", "")) in {"beatsff", "beatsflyingfish"}
        for item in raw_rows
    )
    if existing_manual:
        # Se a ação já existir na planilha, forçamos sua janela para a primeira quinzena
        # e mantemos o limite especial de 3 combos por CPF.
        beats_mapping = resolve_product("Beats Sabores LT 269ml", "MATCH", lookup) or {}
        ff_mapping = resolve_product("Flying Fish LT 350ml", "MATCH", lookup) or {}
        for item in raw_rows:
            if item["categoria"] == "MATCH" and norm(item.get("produto_origem", "")) in {"beatsff", "beatsflyingfish"}:
                item["inicio"], item["fim"] = chosen_window
                item["quinzena"] = category_quarter(*chosen_window)
                item["max_combo_cpf"] = "3"
                item["mecanica"] = ""
                item["tipo_mecanica"] = "Cross sell"
                item["quantidade_mecanica"] = "2"
                item["rotulo_ativacao"] = "Ganhe"
                item["valor_ativacao"] = "1 FF"
                item["mecanica_visual"] = "Leve 2 Beats / Ganhe 1 FF"
                item["produto_origem"] = "BEATS - FF"
                item["nome"] = "BEATS - FF"
                item["nome_abreviado"] = "BEATS - FF"
                item["nome_imagem"] = beats_mapping.get("nome_imagem", item.get("nome_imagem", ""))
                item["caminho_imagem"] = beats_mapping.get("caminho_imagem", item.get("caminho_imagem", ""))
                item["nome_imagem_secundaria"] = ff_mapping.get("nome_imagem", item.get("nome_imagem_secundaria", ""))
                item["caminho_imagem_secundaria"] = ff_mapping.get("caminho_imagem", item.get("caminho_imagem_secundaria", ""))
                item["oferta_manual"] = "sim"
                item["mapeado"] = "sim"
    else:
        beats_mapping = resolve_product("Beats Sabores LT 269ml", "MATCH", lookup)
        ff_mapping = resolve_product("Flying Fish LT 350ml", "MATCH", lookup)
        beats_mapping = beats_mapping or {}
        ff_mapping = ff_mapping or {}
        custom_start, custom_end = chosen_window
        custom_start_ts = timestamp(custom_start)
        raw_rows.append({
            "categoria": "MATCH", "mes": month, "ano": str(year), "inicio": custom_start, "fim": custom_end,
            "quinzena": category_quarter(custom_start, custom_end), "produto_origem": "BEATS - FF",
            "mecanica": "", "tipo_mecanica": "Cross sell", "quantidade_mecanica": "2",
            "rotulo_ativacao": "Ganhe", "valor_ativacao": "1 FF", "max_combo_cpf": "3",
            "mecanica_visual": "Leve 2 Beats / Ganhe 1 FF", "divisao": "", "sinalizador_variacao": "",
            "varejista": "", "abrangencia": "", "tipo_de_para": "MATCH", "promax": "", "eg": "",
            "nome": "BEATS - FF", "nome_abreviado": "BEATS - FF",
            "nome_imagem": beats_mapping.get("nome_imagem", ""),
            "caminho_imagem": beats_mapping.get("caminho_imagem", ""),
            "nome_imagem_secundaria": ff_mapping.get("nome_imagem", ""),
            "caminho_imagem_secundaria": ff_mapping.get("caminho_imagem", ""),
            "oferta_manual": "sim", "mapeado": "sim",
        })

    if not raw_rows:
        warnings.append("Nenhuma oferta válida foi encontrada para o mês selecionado. Confirme que as datas estão nas colunas físicas especificadas e que o mês configurado é o esperado.")

    # TOTAL pode ser um registro-resumo junto com linhas individuais por divisão.
    # Quando houver uma linha individual idêntica, eliminamos a TOTAL para não duplicar o card.
    total_scopes = {"total", "totalgeral", "todadivisao", "todadivisoes"}
    def offer_signature(item: dict[str, Any]) -> tuple[str, ...]:
        fields = ("categoria", "inicio", "fim", "produto_origem", "mecanica", "tipo_mecanica",
                  "quantidade_mecanica", "rotulo_ativacao", "valor_ativacao", "max_combo_cpf", "varejista")
        return tuple(norm(item.get(field, "")) for field in fields)
    individual_signatures = {offer_signature(item) for item in raw_rows if norm(item.get("divisao", "")) not in total_scopes}
    raw_rows = [item for item in raw_rows if not (norm(item.get("divisao", "")) in total_scopes and offer_signature(item) in individual_signatures)]
    for item in raw_rows:
        if norm(item.get("divisao", "")) in total_scopes:
            item["divisao"] = "Todas as divisões"

    # Merge genuinely identical offers across DIVISÃO/GEO values into one card.
    core_fields = ("categoria", "inicio", "fim", "produto_origem", "mecanica", "tipo_mecanica", "quantidade_mecanica", "rotulo_ativacao", "valor_ativacao", "max_combo_cpf", "mecanica_visual", "varejista", "promax", "nome_abreviado")
    merged: dict[tuple[str, ...], dict[str, Any]] = {}
    divisions_for_key: dict[tuple[str, ...], list[str]] = defaultdict(list)
    scopes_for_key: dict[tuple[str, ...], list[str]] = defaultdict(list)
    for row in raw_rows:
        key = tuple(norm(row.get(field, "")) for field in core_fields)
        if key not in merged:
            merged[key] = dict(row)
        division = clean(row.get("divisao", ""))
        if division and division not in divisions_for_key[key]:
            divisions_for_key[key].append(division)
        scope = clean(row.get("abrangencia", ""))
        if scope and scope not in scopes_for_key[key]:
            scopes_for_key[key].append(scope)
    rows = list(merged.values())
    for row in rows:
        key = tuple(norm(row.get(field, "")) for field in core_fields)
        divisions = divisions_for_key.get(key, [])
        scopes = scopes_for_key.get(key, [])
        row["divisao"] = " · ".join(divisions)
        row["abrangencia"] = " · ".join(scopes)
        row["sinalizador_variacao"] = ""

    # Sem abas por UF: consolida mecânicas idênticas e apenas sinaliza produto/período
    # quando existem fórmulas diferentes para divisões distintas.
    variant_groups: dict[tuple[str, ...], list[dict[str, Any]]] = defaultdict(list)
    for row in rows:
        key = tuple(norm(row.get(field, "")) for field in ("categoria", "inicio", "fim", "produto_origem", "varejista"))
        variant_groups[key].append(row)
    for group in variant_groups.values():
        signatures = {tuple(norm(item.get(k, "")) for k in ("mecanica", "tipo_mecanica", "quantidade_mecanica", "rotulo_ativacao", "valor_ativacao", "max_combo_cpf", "mecanica_visual")) for item in group}
        divisions = {
            norm(part)
            for item in group
            for part in clean(item.get("divisao", "")).split(" · ")
            if clean(part)
        }
        scopes = {
            norm(part)
            for item in group
            for part in clean(item.get("abrangencia", "")).split(" · ")
            if clean(part)
        }
        # A custom retailer may have different mechanics by regional suffix in its name.
        # Detect those differences using the parsed suffix even when GEO is the same.
        variation_scopes = scopes if group and group[0].get("categoria") == "PERSONALIZADO" and scopes else divisions
        has_variation = len(signatures) > 1 and len(variation_scopes) > 1
        for item in group:
            if has_variation:
                item["sinalizador_variacao"] = ("Condição varia por região/rede" if item.get("categoria") == "PERSONALIZADO" else "Condição varia por divisão")
                if item.get("categoria") != "PERSONALIZADO":
                    parts = [clean(part) for part in clean(item.get("divisao", "")).split(" · ") if clean(part)]
                    if any(norm(part) == "total" for part in parts):
                        item["divisao"] = "Abrangência geral"
                    else:
                        item["divisao"] = " · ".join(parts)
            elif item.get("categoria") != "PERSONALIZADO":
                # Keep the consolidated divisions so the HTML can show *DF e GO.
                # TOTAL/all-divisions records are suppressed by the renderer.
                item["sinalizador_variacao"] = ""

    frame = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)
    if not frame.empty:
        frame["_start_sort"] = pd.to_datetime(frame["inicio"], errors="coerce")
        frame["_category_sort"] = frame["categoria"].map({c: i for i, c in enumerate(CATEGORIES)})
        frame = frame.sort_values(["_category_sort", "_start_sort", "varejista", "nome_abreviado", "divisao"], kind="stable", na_position="last").drop(columns=["_start_sort", "_category_sort"]).reset_index(drop=True)

    output_file.parent.mkdir(parents=True, exist_ok=True)
    frame.to_csv(output_file, index=False, encoding="utf-8-sig")
    mapping_report = output_file.parent / "mapeamento_produtos_imagens.csv"
    mapping_columns = [
        "categoria", "produto_origem", "nome", "nome_abreviado", "tipo_de_para", "promax", "eg", "abrangencia",
        "nome_imagem", "caminho_imagem", "nome_imagem_secundaria", "caminho_imagem_secundaria", "mapeado"
    ]
    frame[[col for col in mapping_columns if col in frame.columns]].to_csv(
        mapping_report, index=False, encoding="utf-8-sig"
    )
    meta = {
        "mes_atual": month, "ano_atual": year, "titulo": clean(config.get("titulo", "")) or "CALENDÁRIO",
        "logo": clean(config.get("logo", "")) or "logo.png", "fonte_mecanicas": str(offers_path),
        "fonte_de_para": master_description, "apenas_mes_vigente": True, "sem_precos": True,
    }
    meta_file.parent.mkdir(parents=True, exist_ok=True)
    meta_file.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")

    mapped_count = int((frame["mapeado"] == "sim").sum()) if not frame.empty else 0
    categories_counts = frame.groupby("categoria").size().to_dict() if not frame.empty else {}
    missing_map = sorted({(r["categoria"], r["produto_origem"]) for r in rows if r["mapeado"] != "sim"})
    report_lines = [
        "RELATÓRIO DE TRATAMENTO — CALENDÁRIO MERCAPROMO", f"Fonte de ofertas: {offers_path}",
        f"Fonte De/Para: {master_description}", f"Mês de exibição: {month}/{year}",
        f"Arquivo de ligação produto → SKU abreviado → imagem: {mapping_report}",
        "Somente ofertas cujas datas se cruzam com o mês selecionado são incluídas.",
        f"Linhas ignoradas por produto vazio/cabeçalho: {empty_rows}",
        f"Linhas ignoradas por datas vazias/inválidas: {excluded_invalid}",
        f"Linhas excluídas por estarem fora do mês: {excluded_month}",
        f"Cards após agrupamento de ofertas iguais por divisão: {len(frame)} | Mapeados: {mapped_count} | Sem mapeamento: {len(missing_map)}", "",
        "Cards por categoria:",
    ]
    for category in CATEGORIES:
        report_lines.append(f"- {category}: {categories_counts.get(category, 0)}")
    report_lines += ["", "Produtos sem mapeamento (até 100):"]
    report_lines += [f"- {cat}: {product}" for cat, product in missing_map[:100]] or ["- Nenhum."]
    report_lines += ["", "Avisos:"] + [f"- {w}" for w in warnings or ["Nenhum aviso."]]
    report_lines += ["", "Regras aplicadas:",
        "- A fonte principal é Calendario Mercapromo.xlsx; são usadas CERVEJA, NAB, MATCH e PERSONALIZADO; PERSONALIZADO (2) é ignorada.",
        "- CERVEJA/NAB/MATCH: B divisão, C início, D fim, F produto, G mecânica, H tipo, I quantidade/percentual, J label de ativação, K valor de ativação, Q máximo de combos/CPF.",
        "- PERSONALIZADO: instrução na linha 1, cabeçalhos na linha 2 e dados desde a linha 3; A GEO, C varejista, D início, E fim, G produto, H mecânica, I label, J quantidade/percentual, K label, L valor, N máximo por cliente.",
        "- Não há preços nos cards: o destaque principal é a mecânica promocional.",
        "- CERVEJA/NAB/MATCH usam DIVISÃO como etiqueta de abrangência; não são criadas abas por UF.",
        "- Ofertas idênticas em várias DIVISÕES viram um card com as abrangências reunidas.",
        "- Variantes de mecânica no mesmo produto/período podem receber sinalizador de condição por divisão.",
        "- PERSONALIZADO é agrupado pela rede-base; sufixos de UF no nome da rede são convertidos em sinalizadores *DF e GO, etc.",
        "- Todas as ofertas NAB, inclusive Personalizado mapeado como NAB, têm limite de 1 por CPF em todas as quinzenas.",
        "- Todas as ofertas regulares MATCH têm limite de 4 combos por CPF nas duas quinzenas, inclusive no PERSONALIZADO classificado como MATCH pelo De/Para.",
        "- Oferta manual BEATS - FF: Leve 2 Beats / Ganhe 1 FF fica no topo de MATCH na 1ª quinzena, com limite especial de 3 combos por CPF; o gerador a replica no topo da 1ª quinzena de cada rede PERSONALIZADO.",
        "- O De/Para usa Nome para relacionar o nome comercial da promoção, SKU para o nome abreviado exibido e Nome IMG/Caminho para localizar a imagem. O relatório mapeamento_produtos_imagens.csv mostra a ligação por oferta.",
        "- As imagens físicas são localizadas pelo gerador HTML, usando Caminho, Nome IMG e SKU abreviado; relatorio_imagens.txt informa o caminho real localizado.",
    ]
    report_file.parent.mkdir(parents=True, exist_ok=True)
    report_file.write_text("\n".join(report_lines) + "\n", encoding="utf-8")

    print(f"CSV gerado: {output_file}")
    print(f"Relatório gerado: {report_file}")
    print(f"Mês de exibição: {month}/{year}")
    print(f"Total: {len(frame)} | Mapeados: {mapped_count} | Sem mapeamento: {len(missing_map)}")
    for category in CATEGORIES:
        print(f"  {category}: {categories_counts.get(category, 0)} oferta(s)")
    for warning in warnings:
        print(f"Aviso: {warning}")
    return frame


def main() -> None:
    parser = argparse.ArgumentParser(description="Trata somente as mecânicas do mês vigente do Calendário Mercapromo.")
    parser.add_argument("--ofertas", type=Path, default=None, help="Excel principal (padrão: data/unprocessed/Calendario Mercapromo.xlsx)")
    parser.add_argument("--de-para", type=Path, default=None, help="Arquivo opcional De/Para")
    parser.add_argument("--config", type=Path, default=None, help="Excel opcional que possui CONFIG (MesAtual/AnoAtual)")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--report", type=Path, default=DEFAULT_REPORT)
    parser.add_argument("--meta", type=Path, default=DEFAULT_META)
    args = parser.parse_args()
    master = args.de_para if args.de_para is not None else (DEFAULT_MASTER if DEFAULT_MASTER.is_file() else None)
    transform(args.ofertas, master, args.config, args.output, args.report, args.meta)


if __name__ == "__main__":
    main()
