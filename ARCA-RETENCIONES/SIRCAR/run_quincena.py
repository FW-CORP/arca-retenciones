#!/usr/bin/env python3
"""
Entrada rápida: rango de fechas + CUIT agente → TXT SIRCAR 163 con nombre estándar.

Ejecuta `generar_sircar_163_master_dev.py`.

Manual: `Documentacion/MANUAL_ARCA_RETENCIONES_QUINCENA.md`
"""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description="Genera SIRCAR 163 (IIBB) para un rango. Salida en SIRCAR/out/."
    )
    ap.add_argument("--desde", required=True, help="YYYY-MM-DD (incl.)")
    ap.add_argument("--hasta", required=True, help="YYYY-MM-DD (incl.)")
    ap.add_argument(
        "--cuit-agente",
        required=True,
        help="CUIT del agente de retención (11 dígitos, sin guiones).",
    )
    ap.add_argument(
        "--jurisdiccion",
        default="907",
        help="Jurisdicción agente (3 dígitos). Default 907 (Nakel).",
    )
    args, passthrough = ap.parse_known_args(argv)

    root = Path(__file__).resolve().parent
    gen = root / "generar_sircar_163_master_dev.py"
    out = root / "out" / f"SIRCAR_163_{args.desde}_a_{args.hasta}.TXT"

    cmd = [
        sys.executable,
        str(gen),
        "--desde",
        args.desde,
        "--hasta",
        args.hasta,
        "--cuit-agente",
        str(args.cuit_agente),
        "--jurisdiccion",
        str(args.jurisdiccion),
        "--out",
        str(out),
        *passthrough,
    ]
    return int(subprocess.run(cmd).returncode)


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
