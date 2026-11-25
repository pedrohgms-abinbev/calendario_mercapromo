from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"


def run(script: str, args: list[str] | None = None) -> None:
    command = [sys.executable, str(SCRIPTS / script), *(args or [])]
    print(f"\n> {' '.join(command)}", flush=True)
    subprocess.run(command, cwd=ROOT, check=True)


def main() -> None:
    parser = argparse.ArgumentParser(description="Atualiza o Calendário Mercapromo e integra a Black Friday")
    parser.add_argument("--ofertas", type=Path, default=None, help="Excel Calendario Mercapromo.xlsx com abas CERVEJA/NAB/MATCH/PERSONALIZADO")
    parser.add_argument("--de-para", type=Path, default=None, help="De/Para opcional")
    parser.add_argument("--config", type=Path, default=None, help="Excel opcional com a aba CONFIG")
    parser.add_argument(
        "--images-dir", type=Path, action="append", default=[],
        help="Pasta adicional com as imagens (pode repetir); ex.: ../-rvore-de-precos/images/img",
    )
    args = parser.parse_args()

    treat_args: list[str] = []
    for label, value in (("--ofertas", args.ofertas), ("--de-para", args.de_para), ("--config", args.config)):
        if value is not None:
            treat_args += [label, str(value)]

    # Fluxo Mercapromo original: não alteramos seu tratamento.
    run("tratamento.py", treat_args)

    image_args: list[str] = []
    for image_dir in args.images_dir:
        image_args += ["--images-dir", str(image_dir)]
    run("gerar_html.py", image_args)

    # A Black Friday é integrada depois da geração original.
    # O script pula a integração se o JSON enriquecido ainda não existir.
    run("integrar_black_friday.py")

    print(f"\nConcluído. Abra: {ROOT / 'site' / 'calendario_mercapromo.html'}")
    print(f"Relatório Mercapromo: {ROOT / 'data' / 'processed' / 'relatorio_tratamento.txt'}")
    print(f"Relatório Black Friday: {ROOT / 'data' / 'processed' / 'relatorio_black_friday_html.txt'}")


if __name__ == "__main__":
    main()
