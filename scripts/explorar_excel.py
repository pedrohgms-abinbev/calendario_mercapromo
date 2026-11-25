#!/usr/bin/env python3
"""Explora arquivos Excel e produz um resumo seguro para mapear o tratamento.

Uso no repositório:
    python scripts/explorar_excel.py
    python scripts/explorar_excel.py "data/unprocessed/Calendario Mercapromo.xlsx" "cadastro/de_para_mercapromo.xlsx"
    python scripts/explorar_excel.py caminho/arquivo.xlsx --output data/processed/meu_resumo.txt

O relatório mostra nomes reais das abas, posições físicas (A, B, C...), conteúdo
inicial, preenchimento por coluna e exemplos limitados. Não altera os arquivos.
"""
from __future__ import annotations

import argparse
import math
import sys
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pandas as pd

EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xlsb", ".xls"}
IGNORED_DIRS = {".git", ".venv", "venv", "node_modules", "__pycache__"}


def column_letter(number: int) -> str:
    """Converte índice de coluna Excel, começando em 1, para A, B, ..., AA."""
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    try:
        # Cada célula deve ser escalar; bool() também trata numpy.bool_.
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def fmt(value: Any, limit: int = 110) -> str:
    if is_blank(value):
        return ""
    if isinstance(value, (pd.Timestamp, datetime, date)):
        text = value.strftime("%Y-%m-%d %H:%M:%S") if isinstance(value, (pd.Timestamp, datetime)) else value.isoformat()
    elif isinstance(value, float) and math.isfinite(value) and value.is_integer():
        text = str(int(value))
    else:
        text = str(value).replace("\r", " ").replace("\n", " / ").strip()
    if len(text) > limit:
        text = text[: limit - 3] + "..."
    return text


def detect_engine(path: Path) -> str | None:
    suffix = path.suffix.lower()
    if suffix == ".xlsb":
        return "pyxlsb"
    if suffix in {".xlsx", ".xlsm"}:
        return "openpyxl"
    # Para .xls, pandas escolhe o mecanismo disponível (normalmente xlrd).
    return None


def discover_files(root: Path) -> list[Path]:
    """Busca planilhas nas pastas usuais, sem varrer ambientes/cache."""
    folders = [root / "data" / "unprocessed", root / "cadastro"]
    found: set[Path] = set()
    for folder in folders:
        if not folder.exists():
            continue
        for path in folder.rglob("*"):
            if not path.is_file() or path.suffix.lower() not in EXCEL_EXTENSIONS:
                continue
            if path.name.startswith("~$"):
                continue
            if any(part in IGNORED_DIRS for part in path.parts):
                continue
            found.add(path.resolve())
    # Também inclui arquivos Excel diretamente na raiz do projeto.
    if root.exists():
        for path in root.iterdir():
            if path.is_file() and path.suffix.lower() in EXCEL_EXTENSIONS and not path.name.startswith("~$"):
                found.add(path.resolve())
    return sorted(found, key=lambda p: str(p).lower())


def safe_read_sheet(book: pd.ExcelFile, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(book, sheet_name=sheet_name, header=None, dtype=object)


def non_empty_values(series: pd.Series) -> list[Any]:
    return [value for value in series.tolist() if not is_blank(value)]


def unique_examples(values: list[Any], limit: int = 5) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        value_text = fmt(value, 70)
        key = value_text.casefold()
        if not value_text or key in seen:
            continue
        seen.add(key)
        result.append(value_text)
        if len(result) >= limit:
            break
    return result


def summarize_sheet(lines: list[str], sheet_name: str, df: pd.DataFrame, preview_rows: int, sample_count: int) -> None:
    rows, cols = df.shape
    nonempty_rows = int(df.apply(lambda row: any(not is_blank(v) for v in row.tolist()), axis=1).sum()) if rows else 0
    nonempty_cols = int(df.apply(lambda col: any(not is_blank(v) for v in col.tolist()), axis=0).sum()) if cols else 0
    lines.append(f"  ABA: {sheet_name!r}")
    lines.append(f"  Dimensão lida: {rows} linha(s) x {cols} coluna(s); linhas com conteúdo: {nonempty_rows}; colunas com conteúdo: {nonempty_cols}")

    lines.append(f"\n  AMOSTRA DAS PRIMEIRAS {min(preview_rows, rows)} LINHAS (letras são posições físicas do Excel):")
    found_preview = False
    for row_idx in range(min(preview_rows, rows)):
        entries: list[str] = []
        for col_idx in range(cols):
            value = df.iat[row_idx, col_idx]
            if not is_blank(value):
                entries.append(f"{column_letter(col_idx + 1)}={fmt(value, 85)}")
        if entries:
            found_preview = True
            lines.append(f"    Linha Excel {row_idx + 1}: " + " | ".join(entries[:28]))
            if len(entries) > 28:
                lines.append(f"      ... mais {len(entries) - 28} célula(s) preenchida(s) nesta linha")
    if not found_preview:
        lines.append("    (nenhuma célula preenchida nas linhas iniciais)")

    lines.append("\n  RESUMO POR COLUNA:")
    lines.append("    Letra | # | Preenchidas | Vazias | Valores únicos* | Exemplos de valores")
    for col_idx in range(cols):
        values = df.iloc[:, col_idx].tolist()
        filled = non_empty_values(df.iloc[:, col_idx])
        total = len(values)
        unique_keys = {fmt(v, 200).casefold() for v in filled}
        examples = unique_examples(filled, sample_count)
        blanks = total - len(filled)
        letter = column_letter(col_idx + 1)
        example_text = "; ".join(examples) if examples else "(vazia)"
        lines.append(
            f"    {letter:>4} | {col_idx + 1:>2} | {len(filled):>10} | {blanks:>6} | {len(unique_keys):>14} | {example_text}"
        )
    lines.append("    *Valores únicos calculados sobre o texto exibido; datas/números são amostras para orientação.")


def summarize_known_depara(lines: list[str], file_path: Path, frames: dict[str, pd.DataFrame]) -> None:
    """Detecta possíveis duplicidades no De/Para se os nomes dos campos estiverem visíveis."""
    wanted = {"tipo", "promax", "eg", "sku", "nome img", "caminho"}
    for sheet_name, df in frames.items():
        header_info: tuple[int, dict[str, int]] | None = None
        for row_idx in range(min(15, len(df))):
            mapping: dict[str, int] = {}
            for col_idx in range(df.shape[1]):
                value = df.iat[row_idx, col_idx]
                if is_blank(value):
                    continue
                normalized = " ".join(fmt(value, 100).strip().casefold().replace("_", " ").split())
                if normalized in wanted:
                    mapping[normalized] = col_idx
            if "promax" in mapping and ("eg" in mapping or "sku" in mapping):
                header_info = (row_idx, mapping)
                break
        if not header_info:
            continue
        header_row, mapping = header_info
        lines.append("\n  CHECAGEM DE DE/PARA (cabeçalhos detectados automaticamente):")
        lines.append(f"    Arquivo: {file_path.name}; aba: {sheet_name!r}; cabeçalhos na linha Excel {header_row + 1}")
        lines.append("    Campos encontrados: " + ", ".join(f"{name}={column_letter(idx + 1)}" for name, idx in mapping.items()))
        promax_idx = mapping["promax"]
        for alias_name in ("eg", "sku"):
            if alias_name not in mapping:
                continue
            alias_idx = mapping[alias_name]
            alias_to_promax: dict[str, set[str]] = defaultdict(set)
            for row_idx in range(header_row + 1, len(df)):
                alias = df.iat[row_idx, alias_idx]
                promax = df.iat[row_idx, promax_idx]
                if is_blank(alias) or is_blank(promax):
                    continue
                alias_key = fmt(alias, 200).strip().casefold()
                alias_to_promax[alias_key].add(fmt(promax, 100).strip())
            conflicts = [(alias, ids) for alias, ids in alias_to_promax.items() if len(ids) > 1]
            lines.append(f"    Duplicidades de {alias_name.upper()} apontando para Promax diferentes: {len(conflicts)}")
            for alias, ids in conflicts[:20]:
                lines.append(f"      - {alias!r} -> {', '.join(sorted(ids))}")
            if len(conflicts) > 20:
                lines.append(f"      ... mais {len(conflicts) - 20} conflito(s) não exibido(s)")


def explore_file(path: Path, preview_rows: int, sample_count: int) -> tuple[list[str], bool]:
    lines = ["=" * 100, f"ARQUIVO: {path}", f"Nome: {path.name}", f"Extensão: {path.suffix.lower()}"]
    try:
        with pd.ExcelFile(path, engine=detect_engine(path)) as book:
            lines.append("Abas encontradas (nomes exatos):")
            for idx, sheet_name in enumerate(book.sheet_names, start=1):
                lines.append(f"  {idx}. {sheet_name!r}")
            if not book.sheet_names:
                lines.append("  (nenhuma aba identificada)")
                return lines, False
            frames: dict[str, pd.DataFrame] = {}
            for sheet_name in book.sheet_names:
                try:
                    df = safe_read_sheet(book, sheet_name)
                    frames[sheet_name] = df
                    lines.append("")
                    summarize_sheet(lines, sheet_name, df, preview_rows, sample_count)
                except Exception as exc:  # relatório deve continuar para as outras abas
                    lines.append(f"\n  ERRO AO LER A ABA {sheet_name!r}: {type(exc).__name__}: {exc}")
            summarize_known_depara(lines, path, frames)
        return lines, True
    except Exception as exc:
        lines.append(f"ERRO AO ABRIR ARQUIVO: {type(exc).__name__}: {exc}")
        if path.suffix.lower() == ".xls":
            lines.append("Dica: arquivos .xls antigos podem precisar da dependência xlrd: pip install xlrd")
        elif path.suffix.lower() == ".xlsb":
            lines.append("Dica: para .xlsb, instale pyxlsb: pip install pyxlsb")
        return lines, False


def parse_args() -> argparse.Namespace:
    script_root = Path(__file__).resolve().parents[1]
    parser = argparse.ArgumentParser(description="Explora a estrutura e o conteúdo inicial de arquivos Excel.")
    parser.add_argument("files", nargs="*", help="Caminhos dos arquivos Excel. Se omitidos, procura em data/unprocessed e cadastro.")
    parser.add_argument("--output", default=str(script_root / "data" / "processed" / "resumo_exploracao_excel.txt"),
                        help="Caminho do relatório TXT (padrão: data/processed/resumo_exploracao_excel.txt).")
    parser.add_argument("--preview-rows", type=int, default=12, help="Número de linhas iniciais a mostrar por aba (padrão: 12).")
    parser.add_argument("--samples", type=int, default=5, help="Quantidade de exemplos distintos por coluna (padrão: 5).")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    script_root = Path(__file__).resolve().parents[1]
    if args.preview_rows < 1 or args.samples < 1:
        print("Erro: --preview-rows e --samples precisam ser maiores que zero.", file=sys.stderr)
        return 2

    if args.files:
        files = [Path(name).expanduser() for name in args.files]
        files = [p.resolve() if p.is_absolute() else (Path.cwd() / p).resolve() for p in files]
    else:
        files = discover_files(script_root)

    if not files:
        print("Nenhum arquivo Excel encontrado.")
        print("Coloque os arquivos em data/unprocessed/ e cadastro/, ou passe o caminho diretamente:")
        print('  python scripts/explorar_excel.py "data/unprocessed/Calendario Mercapromo.xlsx" "cadastro/de_para_mercapromo.xlsx"')
        return 1

    output_path = Path(args.output).expanduser()
    if not output_path.is_absolute():
        output_path = (Path.cwd() / output_path).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    report: list[str] = [
        "RESUMO DE EXPLORAÇÃO DE EXCEL — CALENDÁRIO MERCAPROMO",
        f"Gerado em: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"Arquivos analisados: {len(files)}",
        "",
        "ATENÇÃO: este relatório contém exemplos limitados de valores reais do arquivo.",
        "Confira se está autorizado compartilhá-lo antes de enviá-lo.",
        "",
        "COMO LER AS COLUNAS:",
        "- As letras A, B, C... são posições físicas do Excel, não os nomes dos cabeçalhos.",
        "- A prévia mostra as primeiras linhas sem promover nenhuma linha a cabeçalho.",
        "- Isso ajuda a descobrir cabeçalhos em linhas 1, 2, 3 etc. ou títulos mesclados.",
        "",
    ]

    successes = 0
    for path in files:
        if not path.exists():
            report.extend(["=" * 100, f"ARQUIVO NÃO ENCONTRADO: {path}", ""])
            continue
        lines, success = explore_file(path, args.preview_rows, args.samples)
        report.extend(lines)
        report.append("")
        if success:
            successes += 1

    report.extend(["=" * 100, f"RESULTADO: {successes}/{len(files)} arquivo(s) aberto(s) com sucesso."])
    output_path.write_text("\n".join(report) + "\n", encoding="utf-8")

    print(f"Arquivos encontrados para análise: {len(files)}")
    print(f"Arquivos abertos com sucesso: {successes}")
    print(f"Relatório gerado: {output_path}")
    print("Envie o conteúdo desse TXT aqui para eu mapear as abas, linhas de cabeçalho e posições das colunas.")
    if successes == 0:
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
