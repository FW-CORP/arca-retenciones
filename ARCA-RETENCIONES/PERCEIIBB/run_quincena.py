#!/usr/bin/env python3
"""
Atajo: rango de fechas + CUIT agente → TXT percepciones IIBB 163 (nombre estándar en PERCEIIBB/out/).

Ejecuta `generar_perceiibb_arca_master_dev.py`.

Manual: `ARCA-RETENCIONES/Documentacion/MANUAL_ARCA_RETENCIONES_QUINCENA.md`
"""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description="Genera TXT SIRCAR percepciones IIBB (163) para un rango. Salida en PERCEIIBB/out/."
    )
    ap.add_argument("--desde", required=True, help="YYYY-MM-DD (incl.)")
    ap.add_argument("--hasta", required=True, help="YYYY-MM-DD (incl.)")
    ap.add_argument(
        "--cuit-agente",
        required=True,
        help="CUIT del agente de percepción (11 dígitos, sin guiones).",
    )
    ap.add_argument(
        "--jurisdiccion",
        default="907",
        help="Jurisdicción agente / venta (3 dígitos). Default 907.",
    )
    args, passthrough = ap.parse_known_args(argv)

    root = Path(__file__).resolve().parent
    gen = root / "generar_perceiibb_arca_master_dev.py"
    out = root / "out" / f"PERCEIIBB_ARCA_{args.desde}_a_{args.hasta}.TXT"

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
