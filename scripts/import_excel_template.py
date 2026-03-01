from __future__ import annotations

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from app import app, import_from_excel_path  # noqa: E402


def main():
    if len(sys.argv) < 2:
        print("Uso: python scripts/import_excel_template.py caminho_da_planilha.xlsx [--reset]")
        raise SystemExit(1)

    xlsx_path = Path(sys.argv[1]).expanduser().resolve()
    reset = "--reset" in sys.argv[2:]

    if not xlsx_path.exists():
        print(f"Arquivo não encontrado: {xlsx_path}")
        raise SystemExit(1)

    with app.app_context():
        counts = import_from_excel_path(xlsx_path, reset=reset)

    print(
        "Importação concluída com sucesso: "
        f"{counts['services']} serviços, "
        f"{counts['collaborators']} colaboradores, "
        f"{counts['attendances']} atendimentos, "
        f"{counts['expenses']} despesas e "
        f"{counts['products']} produtos."
    )


if __name__ == "__main__":
    main()
