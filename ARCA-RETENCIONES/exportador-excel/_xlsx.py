from __future__ import annotations

import subprocess
from pathlib import Path


def convert_csv_to_xlsx(csv_path: Path, xlsx_out: Path) -> Path:
    """
    Convierte CSV → XLSX usando LibreOffice headless.

    Motivación: no asumimos `openpyxl` instalado en este entorno.
    """
    xlsx_out.parent.mkdir(parents=True, exist_ok=True)
    tmp_dir = xlsx_out.parent
    cmd = [
        "libreoffice",
        "--headless",
        "--nologo",
        "--nolockcheck",
        "--nodefault",
        "--norestore",
        "--convert-to",
        "xlsx",
        "--outdir",
        str(tmp_dir),
        str(csv_path),
    ]
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        raise SystemExit(
            "Fallo la conversión a XLSX con LibreOffice.\n"
            f"STDOUT:\n{r.stdout}\nSTDERR:\n{r.stderr}\n"
        )
    produced = tmp_dir / (csv_path.stem + ".xlsx")
    if not produced.exists():
        raise SystemExit(f"LibreOffice no generó el archivo esperado: {produced}")
    if produced.resolve() != xlsx_out.resolve():
        produced.replace(xlsx_out)
    return xlsx_out

