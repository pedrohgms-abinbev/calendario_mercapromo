from __future__ import annotations

import argparse
import base64
import html
import math
import json
import os
import re
import unicodedata
from datetime import datetime
from io import BytesIO
from pathlib import Path
from pathlib import PureWindowsPath
from typing import Any

import pandas as pd
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_INPUT = ROOT / "data" / "processed" / "calendario_mercapromo.csv"
DEFAULT_OUTPUT = ROOT / "site" / "calendario_mercapromo.html"
DEFAULT_META = ROOT / "data" / "processed" / "calendario_mercapromo_meta.json"
IMAGE_DIR = ROOT / "images" / "img"
DEFAULT_LOGOS = (ROOT / "images" / "logo.png", ROOT / "images" / "img" / "logo.png")
IMAGE_ROOTS = (ROOT / "images", ROOT / "assets", ROOT / "static" / "images", ROOT / "data" / "images")
# Additional image roots can be passed with --images-dir or the environment
# variable MERCAPROMO_IMAGES_DIR. These are useful when images still live in
# the existing Árvore de Preços repository beside this new repository.
EXTRA_IMAGE_ROOTS: list[Path] = []
CATEGORIES = ("CERVEJA", "NAB", "MATCH", "PERSONALIZADO")
CATEGORY_COLORS = {"CERVEJA": "#e8d69c", "NAB": "#9ec97b", "MATCH": "#c6a5d5", "PERSONALIZADO": "#9fb6d2"}
MONTHS = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
IMAGE_URI_CACHE: dict[str, str] = {}
IMAGE_FILE_INDEX: dict[str, list[Path]] | None = None
IMAGE_REQUESTS: set[str] = set()
IMAGE_RESOLVED: set[str] = set()
IMAGE_RESOLUTIONS: dict[str, Path] = {}
IMAGE_SEARCH_ROOT_CACHE: list[Path] | None = None
CARD_IMAGE_LINKS: list[dict[str, str]] = []


def clean(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    text = str(value).strip()
    return "" if text.casefold() in {"nan", "none", "null", "nat", "-"} else text


def norm(value: Any) -> str:
    text = unicodedata.normalize("NFKD", clean(value))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]+", "", text.casefold())


def esc(value: Any) -> str:
    return html.escape(clean(value), quote=True)


def date_value(value: Any) -> pd.Timestamp | None:
    text = clean(value)
    if not text:
        return None
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}(?:[ T].*)?", text):
        parsed = pd.to_datetime(text, errors="coerce")
    else:
        parsed = pd.to_datetime(text, dayfirst=True, errors="coerce")
    return None if pd.isna(parsed) else pd.Timestamp(parsed)


def compact_date(value: Any) -> str:
    parsed = date_value(value)
    return parsed.strftime("%d/%m") if parsed is not None else (clean(value) or "—")


def image_search_roots() -> list[Path]:
    """Find the image bank in this repo and in other repositories in the Codespaces workspace."""
    global IMAGE_SEARCH_ROOT_CACHE
    if IMAGE_SEARCH_ROOT_CACHE is not None:
        return IMAGE_SEARCH_ROOT_CACHE

    roots: list[Path] = []
    roots.extend(EXTRA_IMAGE_ROOTS)
    roots.extend([IMAGE_DIR, *IMAGE_ROOTS, ROOT / "img"])

    workspace = ROOT.parent
    # Add conventional folders in sibling repositories first, preserving the old
    # price-tree bank location when both projects are open in the same Codespace.
    if workspace.is_dir():
        try:
            siblings = sorted(
                (p for p in workspace.iterdir() if p.is_dir() and p.resolve() != ROOT.resolve()),
                key=lambda p: p.name.casefold(),
            )
        except OSError:
            siblings = []
        for sibling in siblings:
            for rel in (
                Path("images") / "img", Path("images"), Path("img"), Path("imagens"),
                Path("assets") / "images", Path("static") / "images", Path("data") / "images",
                Path("public") / "images", Path("public") / "img",
            ):
                roots.append(sibling / rel)

        # The bank has lived in different subfolders across project copies. Discover
        # folders named images/img recursively instead of assuming only one repo layout.
        ignored_dirs = {
            ".git", ".venv", "venv", "node_modules", "__pycache__", ".pytest_cache",
            ".mypy_cache", ".next", "dist", "build", "site-packages", ".cache",
        }
        allowed_names = {"img", "imgs", "image", "images", "imagens", "productimages", "bancodeimagens", "bancoimagens"}
        try:
            for current, dirs, _ in os.walk(workspace):
                dirs[:] = [d for d in dirs if d.casefold() not in ignored_dirs]
                current_path = Path(current)
                if current_path.resolve() == workspace.resolve():
                    continue
                if norm(current_path.name) in allowed_names:
                    roots.append(current_path)
        except OSError:
            pass

    # Environment override can specify one or several paths separated by os.pathsep.
    env_value = os.environ.get("MERCAPROMO_IMAGES_DIR", "").strip()
    if env_value:
        for item in env_value.split(os.pathsep):
            if item.strip():
                path = Path(item.strip()).expanduser()
                roots.append(path if path.is_absolute() else ROOT / path)

    unique: list[Path] = []
    seen: set[str] = set()
    for root in roots:
        try:
            key = str(root.resolve()).casefold()
        except OSError:
            key = str(root).casefold()
        if key not in seen and root.is_dir():
            seen.add(key)
            unique.append(root.resolve())
    IMAGE_SEARCH_ROOT_CACHE = unique
    return unique

def image_file_index() -> dict[str, list[Path]]:
    """Index image names/stems across the current and neighboring project folders.

    De/Para cells may contain a Windows path, a relative path, ``AP269.png``
    or just ``AP269``. The folder itself can still belong to the older price-tree
    repository; no copying is needed as long as both repositories exist in the
    same Codespaces workspace when this HTML is generated.
    """
    global IMAGE_FILE_INDEX
    if IMAGE_FILE_INDEX is not None:
        return IMAGE_FILE_INDEX

    index: dict[str, list[Path]] = {}
    supported = {".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp", ".svg", ".tif", ".tiff", ".avif"}
    ignored_dirs = {".git", ".venv", "venv", "node_modules", "__pycache__", ".pytest_cache", ".mypy_cache"}
    seen: set[str] = set()

    roots = [*image_search_roots(), ROOT]
    for root in roots:
        if not root.is_dir():
            continue
        for current, dirs, filenames in os.walk(root):
            dirs[:] = [d for d in dirs if d.casefold() not in ignored_dirs]
            current_path = Path(current)
            for filename in filenames:
                path = current_path / filename
                if path.suffix.lower() not in supported:
                    continue
                try:
                    path_key = str(path.resolve()).casefold()
                except OSError:
                    path_key = str(path).casefold()
                if path_key in seen:
                    continue
                seen.add(path_key)
                for key in {norm(path.name), norm(path.stem)}:
                    if key:
                        index.setdefault(key, []).append(path)

    ordered_roots = [p.resolve() for p in image_search_roots() if p.exists()]
    explicit_roots = [p.resolve() for p in EXTRA_IMAGE_ROOTS if p.exists()]
    def path_priority(path: Path) -> tuple[int, int, int, str]:
        try:
            resolved = path.resolve()
        except OSError:
            resolved = path
        if resolved == IMAGE_DIR.resolve() or IMAGE_DIR.resolve() in resolved.parents:
            priority = 0
        elif any(resolved == root or root in resolved.parents for root in explicit_roots):
            priority = 1
        else:
            matching = [i for i, root in enumerate(ordered_roots) if resolved == root or root in resolved.parents]
            priority = 2 + (min(matching) if matching else len(ordered_roots) + 1)
        return (priority, len(path.parts), len(str(path)), str(path).casefold())

    for key in index:
        index[key].sort(key=path_priority)
    IMAGE_FILE_INDEX = index
    return index


def find_image_file(filename: str) -> Path | None:
    """Resolve an image from its path, name, stem, or abbreviated SKU."""
    raw = clean(filename).strip().strip('"').strip("'")
    if not raw:
        return None
    IMAGE_REQUESTS.add(raw)

    # Excel paths often use Windows separators even when running in Linux.
    portable = raw.replace("\\", "/")
    raw_path = Path(portable).expanduser()
    try:
        win_name = PureWindowsPath(raw).name
    except (TypeError, ValueError):
        win_name = raw_path.name
    filename_only = win_name or raw_path.name or raw
    candidates: list[Path] = []
    if raw_path.is_absolute():
        candidates.append(raw_path)
    else:
        candidates.extend([
            ROOT / raw_path,
            ROOT / "images" / raw_path,
            IMAGE_DIR / raw_path,
            IMAGE_DIR / filename_only,
            ROOT / "assets" / filename_only,
            ROOT / "static" / "images" / filename_only,
            ROOT / "data" / "images" / filename_only,
        ])
        # Search the explicit path override and adjacent repos directly too.
        candidates.extend(root / filename_only for root in image_search_roots())
    target = next((path for path in candidates if path.is_file()), None)
    if target:
        IMAGE_RESOLVED.add(raw)
        IMAGE_RESOLUTIONS[raw] = target.resolve()
        return target

    index = image_file_index()
    filename_variants = [filename_only, Path(filename_only).stem]
    for variant in filename_variants:
        key = norm(variant)
        if not key:
            continue
        paths = index.get(key, [])
        if paths:
            IMAGE_RESOLVED.add(raw)
            IMAGE_RESOLUTIONS[raw] = paths[0].resolve()
            return paths[0]
    return None


def image_data_uri(filename: str) -> str:
    target = find_image_file(filename)
    if target is None:
        return ""
    cache_key = str(target.resolve())
    if cache_key in IMAGE_URI_CACHE:
        return IMAGE_URI_CACHE[cache_key]
    try:
        with Image.open(target) as source:
            im = source.convert("RGBA")
            im.thumbnail((420, 420), Image.Resampling.LANCZOS)
        buffer = BytesIO()
        im.save(buffer, format="WEBP", quality=86, method=2)
        uri = "data:image/webp;base64," + base64.b64encode(buffer.getvalue()).decode("ascii")
    except (OSError, ValueError):
        try:
            ext = target.suffix.lower()
            mime = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".gif": "image/gif", ".webp": "image/webp", ".svg": "image/svg+xml", ".bmp": "image/bmp", ".tif": "image/tiff", ".tiff": "image/tiff", ".avif": "image/avif"}.get(ext)
            if not mime:
                return ""
            uri = f"data:{mime};base64," + base64.b64encode(target.read_bytes()).decode("ascii")
        except OSError:
            return ""
    IMAGE_URI_CACHE[cache_key] = uri
    return uri


def logo_html(logo_name: str = "logo.png") -> str:
    candidates: list[Path] = []
    if logo_name:
        p = Path(logo_name).expanduser()
        candidates.extend([p] if p.is_absolute() else [ROOT / p, ROOT / "images" / p, IMAGE_DIR / p.name])
        found = find_image_file(logo_name)
        if found:
            candidates.append(found)
    candidates.extend(DEFAULT_LOGOS)
    for root in image_search_roots():
        candidates.append(root / Path(logo_name or "logo.png").name)
    seen: set[str] = set()
    for path in candidates:
        key = str(path).casefold()
        if key in seen:
            continue
        seen.add(key)
        if path.is_file():
            uri = image_data_uri(str(path))
            if uri:
                return f'<img class="logo-image" src="{uri}" alt="Logo Ambev">'
    return '<div class="brand-wordmark">ambev</div>'


def mechanic_lines(row: dict[str, str]) -> list[str]:
    visual = clean(row.get("mecanica_visual", ""))
    if visual:
        # The treatment writes a slash-delimited primary/secondary mechanic.
        parts = [x.strip() for x in re.split(r"\s*/\s*", visual) if x.strip()]
        return parts or [visual]
    return [clean(row.get("mecanica", "")) or "Mecânica não informada"]


def build_card(row: dict[str, str], index: int) -> str:
    product = clean(row.get("produto_origem", "Produto"))
    short_name = clean(row.get("nome_abreviado", ""))
    image_name_hint = clean(row.get("nome_imagem", ""))
    is_manual = norm(row.get("oferta_manual", "")) == "sim"
    category = clean(row.get("categoria", ""))
    de_para_type = norm(row.get("tipo_de_para", ""))
    product_identity = norm(" ".join([
        product, clean(row.get("nome", "")), clean(row.get("nome_abreviado", ""))
    ]))
    nab_markers = ("guarana", "pepsi", "gatorade", "fusion", "h2oh", "barep", "aguamineral", "tonica")
    is_nab_offer = (
        category.upper() == "NAB"
        or (category.upper() == "PERSONALIZADO" and (
            de_para_type in {"nab", "naoalcoolica", "bebidanaoalcoolica"}
            or any(marker in product_identity for marker in nab_markers)
        ))
    )
    if not is_manual and (not short_name or norm(short_name) == norm(product)):
        short_name = image_name_hint or short_name
    short_name = short_name or product
    full_name = clean(row.get("nome", "")) or product
    promax = clean(row.get("promax", ""))
    eg = clean(row.get("eg", ""))
    mapped = norm(row.get("mapeado", "")) == "sim"
    retailer = clean(row.get("varejista", ""))
    scope = clean(row.get("abrangencia", ""))
    division = clean(row.get("divisao", ""))
    flag = clean(row.get("sinalizador_variacao", ""))

    def resolve_card_image(*sources: str) -> tuple[str, str, str]:
        for source in sources:
            source = clean(source)
            if source:
                uri = image_data_uri(source)
                if uri:
                    path = IMAGE_RESOLUTIONS.get(source)
                    return uri, (str(path) if path else ""), source
        return "", "", ""

    image, image_path, image_key = resolve_card_image(
        clean(row.get("caminho_imagem", "")), clean(row.get("nome_imagem", "")), short_name
    )
    secondary_image, secondary_image_path, secondary_image_key = resolve_card_image(
        clean(row.get("caminho_imagem_secundaria", "")),
        clean(row.get("nome_imagem_secundaria", "")),
        clean(row.get("nome_abreviado_secundario", "")),
    )
    CARD_IMAGE_LINKS.append({
        "categoria": category, "produto_promocao": product, "sku_abreviado": short_name,
        "nome_img_de_para": image_name_hint, "caminho_de_para": clean(row.get("caminho_imagem", "")),
        "chave_usada": image_key, "arquivo_localizado": image_path,
        "imagem_encontrada": "sim" if image else "não",
        "imagem_secundaria_chave": secondary_image_key, "imagem_secundaria_localizada": secondary_image_path,
    })
    if image and secondary_image:
        image_html = (
            f'<div class="dual-image-pair"><img src="{image}" alt="Beats">'
            f'<span class="image-arrow">+</span><img src="{secondary_image}" alt="Flying Fish"></div>'
        )
    elif image:
        image_html = f'<img src="{image}" alt="{esc(short_name)}">'
    elif secondary_image:
        image_html = f'<img src="{secondary_image}" alt="{esc(short_name)}">'
    else:
        image_html = '<div class="image-placeholder">Imagem indisponível</div>'

    lines = mechanic_lines(row)
    mechanic_kind = norm(row.get("tipo_mecanica", ""))
    discount_text = " ".join(lines)
    is_discount = mechanic_kind.startswith("desconto") or bool(re.search(r"desconto\s*\d*[,.]?\d*\s*%", discount_text, flags=re.I))

    if is_discount:
        # Option 3: highlight the discount number, then show activation threshold and
        # optional per-CPF cap in compact secondary lines. Avoid long slash-delimited text.
        pct_match = re.search(r"(\d+(?:[,.]\d+)?)\s*%", discount_text)
        percent = pct_match.group(1).replace(".", ",") if pct_match else ""
        if not percent:
            raw_percent = clean(row.get("quantidade_mecanica", ""))
            numeric_match = re.search(r"\d+(?:[,.]\d+)?", raw_percent)
            if numeric_match:
                try:
                    numeric = float(numeric_match.group(0).replace(",", "."))
                    if 0 < numeric < 1:
                        numeric *= 100
                    percent = (str(int(numeric)) if numeric.is_integer() else str(numeric).replace(".", ","))
                except ValueError:
                    percent = numeric_match.group(0).replace(".", ",")
        threshold_text = ""
        threshold_match = re.search(r"(?:mínimo\s+de|a partir de|\bmin(?:imo)?\b)[^0-9]*(\d+(?:[,.]\d+)?)", discount_text, flags=re.I)
        threshold = clean(row.get("valor_ativacao", ""))
        if threshold_match:
            threshold = threshold_match.group(1)
        if not threshold:
            threshold_match = re.search(r"/(.*)", discount_text)
            if threshold_match:
                value_match = re.search(r"\d+(?:[,.]\d+)?", threshold_match.group(1))
                threshold = value_match.group(0) if value_match else ""
        if threshold:
            try:
                n = float(threshold.replace(",", "."))
                threshold_label = str(int(n)) if n.is_integer() else str(n).replace(".", ",")
            except ValueError:
                threshold_label = threshold
            unit = "unidade" if threshold_label in {"1", "1,0"} else "unidades"
            threshold_text = f"A partir de {esc(threshold_label)} {unit}"
        maximum = "1" if is_nab_offer else clean(row.get("max_combo_cpf", ""))
        if maximum and maximum not in {"0", "0.0", "-"}:
            if is_nab_offer:
                limit_text = "Máx.: 1 combo por CPF"
            else:
                limit_text = f"Máx. {maximum} por CPF"
            limit_html = f'<div class="discount-limit">{esc(limit_text)}</div>'
        else:
            limit_html = ""
        headline_html = (
            f'<div class="discount-headline">{esc(percent)}<span>% OFF</span></div>'
            if percent else '<div class="discount-headline">DESCONTO</div>'
        )
        promo_html = (
            '<div class="discount-kicker">DESCONTO</div>'
            + headline_html
            + (f'<div class="discount-threshold">{threshold_text}</div>' if threshold_text else "")
            + limit_html
        )
        promo_box_class = "promo-box discount-promo"
    else:
        promo_html = "".join(f'<span class="promo-line">{esc(line)}</span>' for line in lines if line)
        if is_manual:
            maximum = clean(row.get("max_combo_cpf", ""))
            if maximum and maximum not in {"0", "0.0", "-"}:
                promo_html += f'<div class="manual-limit">Máx.: {esc(maximum)} combos por CPF</div>'
        elif is_nab_offer:
            # Explicit on every NAB card, not only in the category footer.
            promo_html += '<div class="manual-limit nab-limit">Máx.: 1 combo por CPF</div>'
        promo_box_class = "promo-box"
    description = clean(row.get("mecanica", ""))
    description_norm = norm(description)
    kind_norm = norm(row.get("tipo_mecanica", ""))
    generic_descriptions = {kind_norm, "leve", "desconto", "leveepague", "descontopercentual", "leveepaguepromocional"}
    description_is_redundant = (
        is_discount
        or description_norm in generic_descriptions
        or description_norm in {norm(x) for x in lines}
    )
    desc_html = f'<div class="mechanic-description">{esc(description)}</div>' if description and not description_is_redundant else ""
    retailer_html = f'<div class="network-label">{esc(retailer)}</div>' if retailer and category != "PERSONALIZADO" else ""
    # The same asterisk convention is used for main and personalized cards.
    # TOTAL / all-divisions rows intentionally have no regional asterisk.
    total_scope_tokens = {"total", "totalgeral", "todadivisao", "todadivisoes", "todasasdivisoes", "abrangenciageral"}
    if category == "PERSONALIZADO":
        badge_scope = scope
    else:
        division_parts = [clean(part) for part in re.split(r"\s*[·;|]\s*", division) if clean(part)]
        division_parts = [part for part in division_parts if norm(part) not in total_scope_tokens]
        badge_scope = " e ".join(division_parts)
    context_html = f'<div class="scope-label">*{esc(badge_scope)}</div>' if badge_scope else ""
    flag_html = f'<div class="difference-flag">⚑ {esc(flag)}</div>' if flag else ""
    map_html = "" if mapped or is_manual else '<div class="mapping-warning">Conferir cadastro da imagem</div>'
    search = norm(" ".join([short_name, full_name, product, promax, eg, retailer, scope, division, description, " ".join(lines)]))
    long_free_text = is_manual or (not clean(row.get("tipo_mecanica", "")) and len(clean(row.get("mecanica_visual", ""))) > 28)
    tile_class = "product-tile special-mechanic" if long_free_text else "product-tile"
    return (
        f'<article class="{tile_class}" data-search="{esc(search)}" data-index="{index}">'
        f'<div class="tile-image">{image_html}</div><div class="tile-body">{retailer_html}'
        f'<h3 title="{esc(full_name)}">{esc(short_name)}</h3>{desc_html}'
        f'<div class="{promo_box_class}">{promo_html}</div>{context_html}{flag_html}{map_html}</div></article>'
    )


def load_meta(path: Path) -> dict[str, Any]:
    if not path.is_file():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {}


def group_periods(subset: pd.DataFrame, category: str) -> list[tuple[str, pd.DataFrame]]:
    parts: list[tuple[str, pd.DataFrame]] = []
    if subset.empty:
        return parts
    work = subset.copy()
    work["_start"] = pd.to_datetime(work["inicio"], errors="coerce")
    work["_end"] = pd.to_datetime(work["fim"], errors="coerce")
    sort_cols = ["_start", "_end", "nome_abreviado"]
    if category == "PERSONALIZADO" and "varejista" in work.columns:
        sort_cols.insert(0, "varejista")
    work = work.sort_values(sort_cols, kind="stable", na_position="last")
    keys = ["inicio", "fim", "quinzena"]
    if category == "PERSONALIZADO":
        keys = ["varejista", *keys]
    for vals, group in work.groupby(keys, sort=False, dropna=False):
        vals = vals if isinstance(vals, tuple) else (vals,)
        if category == "PERSONALIZADO":
            retailer, start, end, quarter = [clean(v) for v in vals]
            title_prefix = f"{retailer} · " if retailer else "Rede não identificada · "
        else:
            start, end, quarter = [clean(v) for v in vals]
            title_prefix = ""
        label_quarter = "1ª Quinzena" if norm(quarter).startswith("1a") else "2ª Quinzena" if norm(quarter).startswith("2a") else (quarter or "Período")
        if start and end:
            heading = f"{title_prefix}{label_quarter}: {compact_date(start)} – {compact_date(end)}"
        else:
            heading = title_prefix + label_quarter
        parts.append((heading, group.drop(columns=["_start", "_end"])))
    return parts


def group_max_combo_note(group: pd.DataFrame) -> str:
    if "max_combo_cpf" not in group.columns or group.empty:
        return ""

    # Determine the displayed category. Personalized rows keep categoria=PERSONALIZADO
    # but have categoria_visual_personalizado set to CERVEJA/NAB/MATCH in the HTML layer.
    visual_categories: set[str] = set()
    for _, row in group.iterrows():
        value = clean(row.get("categoria_visual_personalizado", "")) or clean(row.get("categoria", ""))
        value_norm = norm(value)
        aliases = {"match": "match", "bebidamista": "match", "nab": "nab", "cerv": "cerveja", "cerveja": "cerveja"}
        visual_categories.add(aliases.get(value_norm, value_norm))

    is_match_panel = visual_categories == {"match"}
    if is_match_panel:
        has_beats_ff = any(
            norm(row.get("produto_origem", "")) in {"beatsff", "beatsflyingfish"}
            or norm(row.get("oferta_manual", "")) == "sim"
            for _, row in group.iterrows()
        )
        values = [clean(row.get("max_combo_cpf", "")) for _, row in group.iterrows()]
        if has_beats_ff:
            return ('<div class="period-footer">Máx.: 4 combos por CPF para todos os itens '
                    '· BEATS - FF: Máx.: 3 combos por CPF</div>')
        if values and all(norm(value) in {"4", "4,0", "4.0"} for value in values):
            return '<div class="period-footer">Máx.: 4 combos por CPF para todos os itens</div>'

    values = [(clean(row.get("nome_abreviado", "")) or clean(row.get("produto_origem", "produto")), clean(row.get("max_combo_cpf", "")))
              for _, row in group.iterrows()]
    values = [(name, value) for name, value in values if value]
    if not values:
        return ""
    unique_values = list(dict.fromkeys(value for _, value in values))
    all_rows_have_values = len(values) == len(group)
    if len(unique_values) == 1:
        maximum = unique_values[0]
        unit = "combo" if norm(maximum) in {"1", "1,0", "1.0"} else "combos"
        suffix = " para todos os itens" if all_rows_have_values else " para os itens informados"
        return f'<div class="period-footer">Máx.: {esc(maximum)} {unit} por CPF{suffix}</div>'
    return '<div class="period-footer">O limite de combos por CPF varia conforme o produto</div>'


def period_label(start: Any, end: Any, quarter: Any) -> str:
    q = norm(quarter)
    if q.startswith("1a"):
        prefix = "1ª Quinzena"
    elif q.startswith("2a"):
        prefix = "2ª Quinzena"
    else:
        prefix = clean(quarter) or "Período promocional"
    if clean(start) and clean(end):
        return f"{prefix}: {compact_date(start)} – {compact_date(end)}"
    return prefix


def group_custom_retailer_periods(group: pd.DataFrame) -> list[tuple[str, pd.DataFrame]]:
    """Group a retailer's customized offers by the actual validity window."""
    if group.empty:
        return []
    work = group.copy()
    work["_start"] = pd.to_datetime(work["inicio"], errors="coerce")
    work["_end"] = pd.to_datetime(work["fim"], errors="coerce")
    work = work.sort_values(["_start", "_end", "nome_abreviado", "produto_origem"], kind="stable", na_position="last")
    result: list[tuple[str, pd.DataFrame]] = []
    for keys, items in work.groupby(["inicio", "fim", "quinzena"], sort=False, dropna=False):
        start, end, quarter = keys
        result.append((period_label(start, end, quarter), items.drop(columns=["_start", "_end"])))
    return result


def responsive_columns(count: int, category: str = "") -> int:
    """Choose columns by product count so small categories do not leave awkward gaps."""
    count = max(1, int(count))
    category = clean(category).upper()
    if category == "CERVEJA":
        if count <= 5:
            return count
        if count <= 12:
            return 4
        return 5
    if category in {"NAB", "MATCH"}:
        # Three items fit on one line; four/five use a balanced 2-column grid.
        if count <= 3:
            return count
        if count <= 5:
            return 2
        if count <= 9:
            return 3
        return 4
    if count <= 3:
        return count
    if count <= 8:
        return 2
    if count <= 15:
        return 3
    return 4


def quarter_key(value: Any, start_value: Any = "") -> str:
    text = norm(value)
    if text.startswith("1a"):
        return "q1"
    if text.startswith("2a"):
        return "q2"
    parsed = date_value(start_value)
    if parsed is not None:
        return "q1" if parsed.day <= 8 else "q2"
    return "q1"


def period_range_label(subset: pd.DataFrame) -> str:
    if subset.empty:
        return ""
    starts = [date_value(value) for value in subset.get("inicio", pd.Series(dtype=str)).tolist()]
    ends = [date_value(value) for value in subset.get("fim", pd.Series(dtype=str)).tolist()]
    starts = [value for value in starts if value is not None]
    ends = [value for value in ends if value is not None]
    if not starts or not ends:
        return ""
    return f"{min(starts).strftime('%d/%m')} – {max(ends).strftime('%d/%m')}"


def render_category(category: str, subset: pd.DataFrame, card_index: int) -> tuple[str, int]:
    color = CATEGORY_COLORS.get(category, "#d9ded9")
    ordered = subset.copy()
    if not ordered.empty and clean(category).upper() == "MATCH":
        # Cross-sell BEATS - FF goes first in the first-quincena MATCH panel,
        # including the replicated MATCH panels within personalized networks.
        is_special = (
            ordered.get("oferta_manual", pd.Series([""] * len(ordered), index=ordered.index)).map(norm).eq("sim")
            | ordered.get("produto_origem", pd.Series([""] * len(ordered), index=ordered.index)).map(norm).isin({"beatsff", "beatsflyingfish"})
        )
        ordered["_special_sort"] = (~is_special).astype(int)
        ordered = ordered.sort_values("_special_sort", kind="stable").drop(columns=["_special_sort"])
    cards = "".join(build_card(row.to_dict(), card_index + i) for i, (_, row) in enumerate(ordered.iterrows()))
    if not cards:
        cards = '<div class="category-empty">Nenhuma oferta nesta categoria para este período.</div>'
    columns = responsive_columns(len(ordered), category)
    return (
        f'<section class="category-box category-{norm(category)}" style="--category-color:{color}">'
        f'<header class="category-banner"><h2>{esc(category)}</h2><span>{len(ordered)} oferta(s)</span></header>'
        f'<div class="offers-grid" style="--offer-columns:{columns}">{cards}</div>{group_max_combo_note(ordered)}</section>',
        card_index + len(ordered),
    )


def personalized_visual_category(row: pd.Series) -> str | None:
    """Return the product-family category for a PERSONALIZADO card from the De/Para type."""
    value = norm(row.get("tipo_de_para", ""))
    aliases = {
        "cerv": "CERVEJA", "cerveja": "CERVEJA", "beer": "CERVEJA",
        "nab": "NAB", "naoalcoolica": "NAB", "bebidanaoalcoolica": "NAB",
        "match": "MATCH", "bebidamista": "MATCH",
    }
    if value in aliases:
        return aliases[value]
    product = norm(" ".join([clean(row.get("produto_origem", "")), clean(row.get("nome", "")), clean(row.get("nome_abreviado", ""))]))
    if any(token in product for token in ("beats", "flyingfish", "mikes")):
        return "MATCH"
    if any(token in product for token in ("guarana", "pepsi", "gatorade", "fusion", "h2oh", "agua")):
        return "NAB"
    return None


def generate(input_file: Path, output_file: Path, meta_file: Path) -> None:
    global IMAGE_FILE_INDEX, IMAGE_SEARCH_ROOT_CACHE
    IMAGE_URI_CACHE.clear()
    IMAGE_FILE_INDEX = None
    IMAGE_SEARCH_ROOT_CACHE = None
    IMAGE_REQUESTS.clear()
    IMAGE_RESOLVED.clear()
    IMAGE_RESOLUTIONS.clear()
    CARD_IMAGE_LINKS.clear()
    if not input_file.is_file():
        raise FileNotFoundError(f"Não encontrei o CSV tratado: {input_file}. Execute python scripts/tratamento.py primeiro.")
    df = pd.read_csv(input_file, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    required = {"categoria", "inicio", "fim", "produto_origem", "nome_abreviado", "mapeado", "mecanica_visual"}
    missing = sorted(required - set(df.columns))
    if missing:
        raise ValueError(f"CSV não corresponde ao novo formato de mecânicas. Colunas ausentes: {', '.join(missing)}. Execute python scripts/atualizar.py.")
    if "abrangencia" not in df.columns:
        df["abrangencia"] = ""
    if "varejista" not in df.columns:
        df["varejista"] = ""
    meta = load_meta(meta_file)
    month = clean(meta.get("mes_atual", "")) or MONTHS[datetime.now().month - 1]
    year = clean(meta.get("ano_atual", "")) or str(datetime.now().year)
    logo = logo_html(clean(meta.get("logo", "logo.png")))
    df = df[df["categoria"].isin(CATEGORIES)].copy()
    df = df[df["mes"].map(norm) == norm(month)].copy() if "mes" in df.columns else df
    df["_quarter_key"] = [quarter_key(q, start) for q, start in zip(df.get("quinzena", pd.Series([""] * len(df))), df["inicio"])]
    total = len(df)
    mapped = int((df["mapeado"].map(norm) == "sim").sum()) if total else 0

    # Main calendar is always visible. Both quinzenas are stacked vertically, matching
    # the personalized layout; only personalized retailers are controlled by filter buttons.
    main_views: list[str] = []
    card_index = 0
    main_categories = ("CERVEJA", "NAB", "MATCH")
    for qkey, label in (("q1", "1ª QUINZENA"), ("q2", "2ª QUINZENA")):
        q_rows = df[(df["_quarter_key"] == qkey) & df["categoria"].isin(main_categories)].copy()
        date_label = period_range_label(q_rows)
        q_label = f"{label} · {date_label}" if date_label else label
        beer_html, card_index = render_category("CERVEJA", q_rows[q_rows["categoria"] == "CERVEJA"], card_index)
        nab_html, card_index = render_category("NAB", q_rows[q_rows["categoria"] == "NAB"], card_index)
        match_html, card_index = render_category("MATCH", q_rows[q_rows["categoria"] == "MATCH"], card_index)
        main_views.append(
            f'<section class="period-panel main-period" data-quarter="{qkey}">'
            f'<h3 class="period-heading">{esc(q_label)}</h3>'
            f'<div class="main-layout">{beer_html}{nab_html}{match_html}</div></section>'
        )

    # Each customized retailer has a single filter; regional scope stays on each offer card.
    personalized = df[df["categoria"] == "PERSONALIZADO"].copy()
    manual_mask = (
        (df.get("oferta_manual", pd.Series([""] * len(df), index=df.index)).map(norm) == "sim")
        | df["produto_origem"].map(norm).isin({"beatsff", "beatsflyingfish"})
    )
    beats_ff_seed = df[manual_mask & (df["categoria"] == "MATCH")].head(1)
    networks = sorted({clean(v) for v in personalized["varejista"].tolist() if clean(v)}, key=lambda v: norm(v))
    network_buttons: list[str] = []
    network_views: list[str] = []
    used_view_ids: set[str] = set()
    for network in networks:
        base_id = f"network-{norm(network) or 'rede'}"
        view_id = base_id
        counter = 2
        while view_id in used_view_ids:
            view_id = f"{base_id}-{counter}"
            counter += 1
        used_view_ids.add(view_id)
        network_buttons.append(
            f'<button type="button" class="filter-button network-filter" data-view-target="{view_id}" '
            f'aria-pressed="false">{esc(network)}</button>'
        )
        n_rows = personalized[personalized["varejista"].map(norm) == norm(network)].copy()
        # The manual "Leve 2 Beats / Ganhe 1 FF" action belongs to the first
        # quinzena in the main calendar AND in every personalized network view.
        # Use each retailer's own first-window dates when available.
        if not beats_ff_seed.empty:
            def is_beats_ff_row(row: pd.Series) -> bool:
                return (
                    norm(row.get("produto_origem", "")) in {"beatsff", "beatsflyingfish"}
                    or norm(row.get("oferta_manual", "")) == "sim"
                )
            if not n_rows.empty:
                n_rows = n_rows[~n_rows.apply(is_beats_ff_row, axis=1)].copy()
            if "categoria_visual_personalizado" not in n_rows.columns:
                n_rows["categoria_visual_personalizado"] = n_rows.apply(personalized_visual_category, axis=1)
            q1_rows = n_rows[n_rows.get("_quarter_key", pd.Series(index=n_rows.index, dtype=str)) == "q1"].copy()
            if not q1_rows.empty:
                q1_windows = q1_rows[["inicio", "fim"]].drop_duplicates().copy()
                q1_windows["_sort_start"] = pd.to_datetime(q1_windows["inicio"], errors="coerce")
                q1_windows = q1_windows.sort_values("_sort_start", kind="stable", na_position="last")
                custom_start, custom_end = str(q1_windows.iloc[0]["inicio"]), str(q1_windows.iloc[0]["fim"])
            else:
                seed = beats_ff_seed.iloc[0]
                custom_start, custom_end = clean(seed.get("inicio", "")), clean(seed.get("fim", ""))
            manual_copy = beats_ff_seed.iloc[0].copy()
            manual_copy["categoria"] = "PERSONALIZADO"
            manual_copy["varejista"] = network
            manual_copy["inicio"] = custom_start
            manual_copy["fim"] = custom_end
            manual_copy["quinzena"] = "1ª QUINZENA"
            manual_copy["_quarter_key"] = "q1"
            manual_copy["abrangencia"] = ""
            manual_copy["divisao"] = ""
            manual_copy["sinalizador_variacao"] = ""
            manual_copy["max_combo_cpf"] = "3"
            manual_copy["tipo_de_para"] = "MATCH"
            manual_copy["categoria_visual_personalizado"] = "MATCH"
            n_rows = pd.concat([n_rows, pd.DataFrame([manual_copy])], ignore_index=True)
        if "categoria_visual_personalizado" not in n_rows.columns:
            n_rows["categoria_visual_personalizado"] = n_rows.apply(personalized_visual_category, axis=1)
        period_sections: list[str] = []
        for qkey, qlabel in (("q1", "1ª QUINZENA"), ("q2", "2ª QUINZENA")):
            q_rows = n_rows[n_rows["_quarter_key"] == qkey].copy()
            date_label = period_range_label(q_rows)
            panel_label = f"{qlabel} · {date_label}" if date_label else qlabel
            category_sections: list[str] = []
            for category in main_categories:
                cat_rows = q_rows[q_rows["categoria_visual_personalizado"] == category].copy()
                cat_html, card_index = render_category(category, cat_rows, card_index)
                category_sections.append(cat_html)
            unknown_rows = q_rows[q_rows["categoria_visual_personalizado"].isna() | (q_rows["categoria_visual_personalizado"] == "")].copy()
            if not unknown_rows.empty:
                unknown_cards = "".join(build_card(row.to_dict(), card_index + i) for i, (_, row) in enumerate(unknown_rows.iterrows()))
                card_index += len(unknown_rows)
                category_sections.append(
                    '<section class="category-box category-unmapped" style="--category-color:#d5d8d9">'
                    f'<header class="category-banner"><h2>REVISAR CADASTRO</h2><span>{len(unknown_rows)} oferta(s)</span></header>'
                    f'<div class="offers-grid">{unknown_cards}</div></section>'
                )
            period_sections.append(
                f'<section class="period-panel network-period"><h3 class="period-heading">{esc(panel_label)}</h3>'
                f'<div class="network-category-layout">{"".join(category_sections)}</div>'
                f'{group_max_combo_note(q_rows)}</section>'
            )
        network_views.append(
            f'<section class="view-panel network-view" data-view-panel="{view_id}" hidden>'
            f'<header class="view-heading network-view-heading"><div><div class="view-eyebrow">PERSONALIZADO</div>'
            f'<h2>{esc(network)}</h2></div><span class="view-period">{esc(month)} {esc(year)}</span></header>'
            f'<p class="scope-explainer">* Indica a abrangência regional da oferta.</p>'
            f'{"".join(period_sections)}</section>'
        )

    main_calendar_id = "main-calendar-view"
    main_calendar_button = (
        f'<button type="button" class="filter-button main-filter active" data-view-target="{main_calendar_id}" '
        'aria-pressed="true">Calendário principal</button>'
    )
    filterbar = (
        '<nav class="filter-shell" aria-label="Filtros do Calendário Mercapromo">'
        '<div class="filter-groups">'
        '<section class="filter-group main-filter-group">'
        '<div class="filter-group-label">CALENDÁRIO PRINCIPAL</div>'
        '<div class="filter-row">' + main_calendar_button + '</div>'
        '</section>'
        + (
            '<section class="filter-group personalized-filter-group filter-only-networks">'
            '<div class="filter-group-label">PERSONALIZADOS · FILTRAR POR REDE</div>'
            '<div class="filter-row">' + "".join(network_buttons) + '</div>'
            '</section>'
            if network_buttons else ''
        )
        + '</div></nav>'
    )
    now = datetime.now().strftime("%d/%m/%Y %H:%M")
    css = r'''
:root{--ink:#141619;--muted:#616970;--line:#b8bdba;--blue:#07519a;--shadow:0 4px 7px #00000020}
*{box-sizing:border-box}html{background:#e3e4e0;scroll-behavior:smooth}body{margin:0;color:var(--ink);font-family:Arial,Helvetica,sans-serif;background:#e3e4e0}
.page{max-width:1900px;margin:0 auto;padding:18px clamp(10px,1.5vw,26px) 24px}
.masthead{display:flex;align-items:center;justify-content:space-between;gap:18px;margin:0 0 12px;flex-wrap:wrap}.brand-title{display:flex;align-items:center;gap:12px;flex-wrap:wrap}.brand-title h1{margin:0;color:#07519a;font-size:clamp(28px,3vw,48px);font-weight:1000;letter-spacing:-2px;line-height:1;text-transform:uppercase}.brand-divider{width:5px;height:39px;background:#07519a;border-radius:2px}.brand-subtitle{font-size:clamp(22px,2.3vw,35px);font-weight:1000;letter-spacing:-.7px;color:#101010;text-transform:uppercase}.brand-wordmark{font-size:clamp(23px,2vw,32px);font-weight:1000;letter-spacing:-1.7px;color:#32363b}.logo-image{max-width:130px;max-height:52px;object-fit:contain}.month-stamp{margin:-2px 0 10px 2px;color:#39424a;font-weight:900;font-size:clamp(13px,1.1vw,17px);text-transform:uppercase}
.filter-shell{position:sticky;top:0;z-index:20;background:rgba(243,244,241,.97);backdrop-filter:blur(8px);border:1px solid #c1c6c2;border-radius:14px;padding:10px 13px 12px;margin:0 0 14px;box-shadow:0 3px 9px #00000012}.filter-groups{display:flex;align-items:flex-start;gap:18px;flex-wrap:wrap}.filter-group{display:flex;flex-direction:column;gap:7px;min-width:max-content}.filter-group-label{font-size:10px;letter-spacing:1.15px;font-weight:1000;color:#626b70;margin-left:2px}.personalized-filter-group{padding-left:16px;border-left:1px solid #c4c8c5}.filter-row{display:flex;align-items:center;gap:7px;flex-wrap:wrap}.filter-button{appearance:none;border:1px solid #abb5bd;background:#fff;color:#20272d;border-radius:9px;padding:10px 15px;min-height:38px;font:900 12px/1 Arial,Helvetica,sans-serif;cursor:pointer;transition:background .16s ease,color .16s ease,border-color .16s ease}.filter-button:hover{border-color:#07519a}.filter-button.active{background:#07519a;border-color:#07519a;color:#fff;box-shadow:0 2px 4px #07519a2b}.main-filter{background:#eaf2fa;border-color:#9dbbd7}.main-filter.active{background:#07519a;border-color:#07519a;color:#fff}.network-filter{background:#f7f6f9;border-color:#b9b2c7}.network-filter.active{background:#5b4777;border-color:#5b4777;color:#fff}.filter-divider{align-self:stretch;width:1px;background:#c4c8c5;margin:2px 5px}.filter-label{font-size:10px;letter-spacing:.8px;color:#68636e;font-weight:1000;padding:0 3px;white-space:nowrap}
.view-panel[hidden]{display:none!important}.view-panel{scroll-margin-top:110px}.network-view{min-height:0}.view-heading{display:flex;align-items:center;justify-content:space-between;gap:12px;margin:3px 2px 12px;padding:0 2px}.view-eyebrow{font-size:10px;letter-spacing:1.7px;color:#5f6a70;font-weight:1000;margin-bottom:4px}.view-heading h2{font-size:clamp(23px,2.2vw,34px);line-height:1;margin:0;font-weight:1000;letter-spacing:-.6px}.view-period{font-size:12px;font-weight:900;color:#5d666b;white-space:nowrap;text-transform:uppercase}.main-layout,.network-category-layout{display:grid;grid-template-columns:minmax(0,1.55fr) minmax(0,.98fr) minmax(0,.90fr);gap:8px;align-items:start;min-height:0}.category-box{display:flex;flex-direction:column;background:linear-gradient(120deg,#f5f5f2 0%,#e8e9e5 100%);border:1px solid #b8bfbc;border-radius:17px;padding:8px 9px 9px;box-shadow:0 3px 5px #00000010;min-width:0;height:auto;align-self:start}.category-banner{display:flex;align-items:center;justify-content:space-between;gap:8px;background:var(--category-color);padding:6px 10px;margin:0 0 7px;border:1px solid #c2c4bf;border-radius:8px}.category-banner h2{margin:0;font-size:clamp(16px,1.4vw,22px);font-weight:1000;text-transform:uppercase}.category-banner span{font-size:10px;font-weight:900;color:#313a3d;white-space:nowrap}.secondary-categories{display:grid;grid-template-columns:1fr 1fr;gap:13px;align-items:start}.offers-grid{display:grid;grid-template-columns:repeat(var(--offer-columns,4),minmax(0,1fr));gap:7px;align-items:stretch;align-content:start}.product-tile{display:grid;grid-template-columns:minmax(58px,32%) minmax(0,1fr);align-items:center;gap:7px;padding:5px 6px;min-width:0;position:relative;background:rgba(255,255,255,.44);border:1px solid #c8cdca;border-radius:11px;min-height:104px}.product-tile.special-mechanic{grid-column:1/-1;grid-template-columns:minmax(80px,24%) minmax(0,1fr)}.category-nab .product-tile,.category-match .product-tile{grid-template-columns:minmax(0,1fr);grid-template-rows:76px auto;align-items:start;gap:3px;padding:7px 6px;min-height:0;text-align:center}.category-nab .tile-image,.category-match .tile-image{height:74px}.category-nab .tile-image img,.category-match .tile-image img{max-height:70px}.category-nab .tile-body,.category-match .tile-body{gap:4px}.category-nab .tile-body h3,.category-match .tile-body h3{font-size:clamp(11px,.82vw,14px);line-height:1.05;text-align:center;overflow-wrap:normal;word-break:normal;hyphens:none}.category-nab .scope-label,.category-match .scope-label{align-self:center}.category-nab .product-tile.special-mechanic,.category-match .product-tile.special-mechanic{grid-column:1/-1;grid-template-columns:minmax(72px,22%) minmax(0,1fr);grid-template-rows:auto;align-items:center;text-align:left}.category-nab .product-tile.special-mechanic .tile-image,.category-match .product-tile.special-mechanic .tile-image{height:86px}.category-nab .product-tile.special-mechanic .tile-body h3,.category-match .product-tile.special-mechanic .tile-body h3{text-align:left}.category-nab .product-tile.special-mechanic .scope-label,.category-match .product-tile.special-mechanic .scope-label{align-self:flex-start}.tile-image{display:flex;align-items:center;justify-content:center;min-width:0;height:100px}.tile-image img{display:block;max-width:100%;max-height:96px;object-fit:contain;filter:drop-shadow(0 2px 2px #00000013)}.dual-image-pair{width:100%;display:flex;align-items:center;justify-content:space-evenly;gap:3px}.dual-image-pair img{max-width:43%;max-height:85px}.image-arrow{font-size:18px;font-weight:1000;color:#07519a}.image-placeholder{width:90%;height:58px;border:1px dashed #c4c7c2;border-radius:8px;display:flex;align-items:center;justify-content:center;text-align:center;font-size:9px;color:#7a7e7a;background:#f6f6f1}
.tile-body{min-width:0;display:flex;flex-direction:column;align-items:stretch;gap:4px}.tile-body h3{font-size:clamp(12px,.98vw,17px);font-weight:1000;line-height:1.08;margin:0;text-transform:uppercase;overflow-wrap:normal;word-break:normal;hyphens:none}.category-tag{align-self:flex-start;font-size:8px;line-height:1;background:var(--category-color);border-radius:4px;padding:3px 5px;font-weight:1000;letter-spacing:.3px}.tag-cerveja{--category-color:#e8d69c}.tag-nab{--category-color:#9ec97b}.tag-match{--category-color:#c6a5d5}.tag-personalizado{--category-color:#9fb6d2}.promo-box{width:100%;background:#e7d1bf;border:1px solid #cbb6a6;border-radius:9px;padding:7px 5px;text-align:center;box-shadow:0 2px 3px #00000013;display:flex;flex-direction:column;justify-content:center;gap:4px;min-height:50px;min-width:0}.promo-line{display:block;font-size:clamp(12px,.95vw,16px);font-weight:1000;line-height:1.12;overflow-wrap:normal;word-break:normal;hyphens:none}.promo-line:first-child{font-size:clamp(13px,1.02vw,17px)}.discount-promo{align-items:center;gap:3px;padding:8px 6px;min-height:86px;overflow:visible;background:#f3ddca;border-color:#c9ad95}.discount-kicker{font-size:9px;font-weight:1000;letter-spacing:1.1px;line-height:1;color:#63442e}.discount-headline{display:flex;justify-content:center;align-items:baseline;gap:4px;font-size:clamp(24px,1.75vw,36px);font-weight:1000;line-height:1;color:#000000;white-space:nowrap;max-width:100%}.discount-headline span{font-size:.46em;letter-spacing:.35px}.discount-threshold{font-size:clamp(10px,.78vw,13px);font-weight:900;line-height:1.15;color:#303a35;white-space:normal;overflow-wrap:normal;word-break:normal}.discount-limit{font-size:10px;font-weight:800;line-height:1.15;color:#5d6260;white-space:normal;overflow-wrap:normal;word-break:normal}.manual-limit{font-size:10px;font-weight:900;line-height:1.15;color:#353535;margin-top:2px}.mechanic-description{font-size:9px;line-height:1.2;color:#41494c;font-weight:700;overflow-wrap:normal;word-break:normal}.scope-label{display:inline-flex;align-self:flex-start;max-width:100%;font-size:10px;font-weight:1000;color:#07519a;line-height:1.15;overflow-wrap:anywhere;background:#e9f1f9;border:1px solid #b6cde1;border-radius:6px;padding:4px 6px;letter-spacing:.1px}.difference-flag{font-size:8px;font-weight:1000;color:#6b4b00;background:#fff0bf;border:1px solid #e7cb76;padding:3px 4px;border-radius:5px;line-height:1.15}.network-label{font-size:9px;font-weight:1000;text-transform:uppercase;color:#07519a;line-height:1.1}.mapping-warning{font-size:8px;font-weight:1000;background:#fff0f0;color:#a52929;padding:3px 5px;border-radius:4px}.category-empty,.empty-state{padding:9px 7px;text-align:center;background:#f7f7f4;border:1px dashed #b7bfbe;border-radius:9px;color:#687077;font-weight:700;font-size:10px}.period-panel{background:linear-gradient(120deg,#f5f4f7 0%,#ebe9f0 100%);border:1px solid #b8b4c1;border-radius:15px;padding:8px;margin:0 0 9px;box-shadow:0 3px 5px #00000012;min-width:0}.main-calendar{margin-bottom:18px}.main-period .period-heading{margin:0 0 6px}.main-calendar-heading,.personalized-heading{margin:3px 2px 8px}.personalized-calendars{margin-top:10px}.filter-only-networks{padding-left:0;border-left:0}.period-heading{display:flex;align-items:center;justify-content:space-between;gap:10px;margin:0 0 7px;font-size:clamp(16px,1.35vw,21px);font-weight:1000;line-height:1.15}.network-offers{grid-template-columns:repeat(4,minmax(0,1fr))}.network-category-layout>.category-box{min-width:0}.category-unmapped{grid-column:1/-1}.scope-explainer{margin:-4px 2px 12px;font-size:11px;font-weight:800;color:#68636e}.period-footer{margin:7px 3px 0;text-align:right;font-size:10px;font-weight:800;color:#666e71;line-height:1.2}.page-footer{display:flex;justify-content:space-between;gap:12px;flex-wrap:wrap;padding:14px 6px;color:#858b8d;font-size:10px}
@media(max-width:1450px){.network-offers{grid-template-columns:repeat(3,minmax(0,1fr))}.product-tile{grid-template-columns:minmax(54px,32%) minmax(0,1fr);gap:6px}.tile-image{height:92px}.tile-image img{max-height:88px}}
@media(max-width:1050px){.main-layout,.network-category-layout{grid-template-columns:minmax(0,1.02fr) minmax(0,.98fr);align-items:start}.category-cerveja{grid-column:1/-1}.network-offers{grid-template-columns:repeat(2,minmax(0,1fr))}.view-heading h2{font-size:25px}}
@media(max-width:760px){.page{padding:12px 8px 20px}.masthead{margin-bottom:10px}.brand-title h1{font-size:28px;letter-spacing:-1px}.brand-divider{height:27px}.brand-subtitle{font-size:20px}.logo-image{max-width:100px;max-height:42px}.filter-shell{position:static;padding:8px;margin-bottom:12px}.filter-groups{gap:10px}.personalized-filter-group{padding-left:0;border-left:0;border-top:1px solid #c4c8c5;padding-top:9px;width:100%}.filter-row{gap:6px}.filter-button{padding:9px 11px;min-height:34px;font-size:10px}.filter-divider{height:24px}.view-heading h2{font-size:22px}.view-period{font-size:10px}.category-box{padding:8px;border-radius:15px}.category-banner{padding:6px 9px}.main-layout,.network-category-layout{grid-template-columns:1fr}.category-cerveja{grid-column:auto}.category-cerveja .offers-grid,.category-nab .offers-grid,.category-match .offers-grid{grid-template-columns:repeat(2,minmax(0,1fr))}.offers-grid{gap:7px}.product-tile{grid-template-columns:minmax(48px,32%) minmax(0,1fr);gap:5px;padding:5px;min-height:104px}.product-tile.special-mechanic{grid-column:1/-1;grid-template-columns:minmax(60px,25%) minmax(0,1fr)}.tile-image{height:82px}.tile-image img{max-height:78px}.tile-body h3{font-size:12px}.promo-box{padding:5px 4px;min-height:44px}.promo-line{font-size:11px}.promo-line:first-child{font-size:12px}.discount-promo{min-height:78px;padding:7px 4px}.discount-headline{font-size:clamp(22px,5.3vw,29px)}.discount-threshold{font-size:11px}.discount-limit{font-size:9px}.network-offers{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media(max-width:440px){.main-layout,.network-category-layout{grid-template-columns:1fr}.category-cerveja .offers-grid,.category-nab .offers-grid,.category-match .offers-grid,.network-offers{grid-template-columns:1fr}.product-tile{grid-template-columns:76px minmax(0,1fr)}.product-tile.special-mechanic{grid-template-columns:68px minmax(0,1fr)}.tile-image{height:80px}.tile-image img{max-height:76px}.view-heading{align-items:flex-start}.view-period{white-space:normal;text-align:right}.discount-headline{font-size:26px}.discount-threshold{font-size:11px}}
@media print{html,body{background:#fff}.page{max-width:none;padding:5mm}.filter-shell{position:static}.filter-groups{display:block}.view-panel[hidden]{display:block!important;break-before:page}.main-layout,.network-category-layout{display:grid;grid-template-columns:1.55fr .95fr .95fr}.category-box,.period-panel{break-inside:avoid;box-shadow:none}.product-tile{break-inside:avoid}.page-footer{display:none}}
'''
    month_label = f"{month} {year}".strip()
    main_calendar_html = (
        f'<section class="main-calendar view-panel" data-view-panel="{main_calendar_id}"><header class="view-heading main-calendar-heading">'
        '<div><div class="view-eyebrow">MECÂNICAS PROMOCIONAIS</div><h2>CALENDÁRIO MERCAPROMO</h2></div>'
        f'<span class="view-period">{esc(month)} {esc(year)}</span></header>'
        + "".join(main_views) + '</section>'
    )
    # Main calendar and personalized network views are siblings controlled by
    # the same filter state: choosing a network hides the main calendar entirely.
    sections_html = main_calendar_html + '<div id="personalized-views">' + "".join(network_views) + '</div>'
    head = (
        '<!doctype html><html lang="pt-BR"><head><meta charset="utf-8">'
        '<meta name="viewport" content="width=device-width,initial-scale=1">'
        f'<title>MERCAPROMO | CALENDÁRIO · {esc(month_label)}</title><style>{css}</style></head><body><main class="page">'
    )
    header = (
        f'<header class="masthead"><div class="brand-title"><h1>MERCAPROMO</h1><span class="brand-divider"></span>'
        f'<span class="brand-subtitle">CALENDÁRIO</span></div>{logo}</header>'
        f'<div class="month-stamp">MÊS VIGENTE · {esc(month_label)}</div>'
        f'{filterbar}<div id="calendar-views">{sections_html}</div>'
        f'<footer class="page-footer"><span>Calendário Mercapromo · Mecânicas promocionais do mês vigente</span><span>* Indica a abrangência regional da oferta · Gerado em {esc(now)}</span></footer>'
        '<script>(function(){const buttons=[...document.querySelectorAll("[data-view-target]")];'
        'const panels=[...document.querySelectorAll("[data-view-panel]")];'
        'function activate(id){panels.forEach(p=>{p.hidden=p.dataset.viewPanel!==id;});'
        'buttons.forEach(b=>{const on=b.dataset.viewTarget===id;b.classList.toggle("active",on);b.setAttribute("aria-pressed",on?"true":"false");});'
        'const active=panels.find(p=>p.dataset.viewPanel===id);if(active){active.scrollIntoView({block:"start",behavior:"smooth"});}}'
        'buttons.forEach(b=>b.addEventListener("click",()=>activate(b.dataset.viewTarget)));})();</script>'
    )
    output_file.parent.mkdir(parents=True, exist_ok=True)
    output_file.write_text(head + header + '</main></body></html>', encoding="utf-8")
    print(f"HTML gerado: {output_file}")
    print(f"Mês exibido: {month_label}")
    print(f"Ofertas: {total} | Mapeadas: {mapped} | Sem mapeamento: {total-mapped}")
    print(f"Ofertas CERVEJA/NAB/MATCH: {int(df['categoria'].isin(main_categories).sum())}")
    print(f"Redes PERSONALIZADO no filtro: {len(networks)}")
    print(f"Imagens únicas incorporadas: {len(IMAGE_URI_CACHE)}")
    image_roots = image_search_roots()
    if IMAGE_REQUESTS and not IMAGE_URI_CACHE:
        print("ATENÇÃO: nenhuma imagem foi incorporada. Confira o relatório de imagens e se o banco está disponível neste Codespace.")
        print('Para indicar a pasta: python scripts/atualizar.py --images-dir "/caminho/real/images/img"')
    print(f"Pastas de imagens pesquisadas: {len(image_roots)}")
    for image_root in image_roots[:20]:
        print(f"  - {image_root}")
    image_missing = IMAGE_REQUESTS - IMAGE_RESOLVED
    card_missing = [item for item in CARD_IMAGE_LINKS if item.get("imagem_encontrada") != "sim"]
    print(f"Cards sem imagem localizada: {len(card_missing)}")
    if card_missing:
        print("Primeiros cards sem imagem: " + ", ".join(item.get("sku_abreviado", "") or item.get("produto_promocao", "") for item in card_missing[:15]))
    print(f"Aliases/caminhos alternativos sem correspondência exata: {len(image_missing)}")
    if image_missing:
        print("Primeiros aliases sem correspondência exata: " + ", ".join(sorted(image_missing)[:15]))

    report_file = ROOT / "data" / "processed" / "relatorio_imagens.txt"
    report_lines = [
        "RELATÓRIO DE IMAGENS — CALENDÁRIO MERCAPROMO",
        f"HTML gerado: {output_file}",
        f"Imagens incorporadas (arquivos únicos): {len(IMAGE_URI_CACHE)}",
        f"Aliases/caminhos tentados: {len(IMAGE_REQUESTS)}",
        f"Aliases localizados: {len(IMAGE_RESOLVED)}",
        "", "Pastas pesquisadas:",
        *[f"- {root}" for root in image_roots],
        "", "Vínculo promoção → SKU abreviado → arquivo de imagem:",
        *[
            f"- [{item['categoria']}] {item['produto_promocao']} -> SKU '{item['sku_abreviado']}' -> "
            f"Nome IMG '{item['nome_img_de_para']}' -> {item['arquivo_localizado'] or 'IMAGEM NÃO ENCONTRADA'}"
            + (f" | secundária: {item['imagem_secundaria_localizada']}" if item['imagem_secundaria_localizada'] else "")
            for item in CARD_IMAGE_LINKS
        ],
        "", "Aliases localizados e arquivo físico:",
        *[f"- {alias} -> {path}" for alias, path in sorted(IMAGE_RESOLUTIONS.items(), key=lambda kv: kv[0].casefold())],
        "", "Cards sem imagem localizada:",
        *([f"- {item['categoria']}: {item['produto_promocao']} (SKU {item['sku_abreviado']})" for item in CARD_IMAGE_LINKS if item.get("imagem_encontrada") != "sim"] if card_missing else ["- Nenhum"]),
        "", "Aliases/caminhos alternativos sem correspondência exata (podem ser tentativas de fallback):",
        *([f"- {alias}" for alias in sorted(image_missing, key=str.casefold)] if image_missing else ["- Nenhum"]),
    ]
    report_file.parent.mkdir(parents=True, exist_ok=True)
    report_file.write_text("\n".join(report_lines) + "\n", encoding="utf-8")
    print(f"Relatório de imagens: {report_file}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Gera o HTML estático no layout de referência Mercapromo, focado em mecânicas.")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT, help="CSV de mecânicas tratado")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT, help="Caminho do HTML")
    parser.add_argument("--meta", type=Path, default=DEFAULT_META, help="JSON de configuração")
    parser.add_argument(
        "--images-dir", type=Path, action="append", default=[],
        help="Pasta de imagens adicional (pode ser repetida), por exemplo ../-rvore-de-precos/images/img",
    )
    args = parser.parse_args()
    for p in args.images_dir:
        resolved = p.expanduser()
        EXTRA_IMAGE_ROOTS.append(resolved if resolved.is_absolute() else (ROOT / resolved).resolve())
    # Print image search folders for transparent diagnostics.
    generate(args.input, args.output, args.meta)


if __name__ == "__main__":
    main()
