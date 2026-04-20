#!/usr/bin/env python3
"""
Entrada rápida: rango de fechas → TXT SICORE (Ganancias) con nombre estándar.

Ejecuta `generar_sicore_v9_retenciones.py` y, por defecto, valida posiciones clave
del archivo generado.

Manual: `Documentacion/MANUAL_ARCA_RETENCIONES_QUINCENA.md`
"""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description="Genera SICORE v9 (retenciones Ganancias) para un rango. Salida en SICORE/out/."
    )
    ap.add_argument("--desde", required=True, help="YYYY-MM-DD (incl.)")
    ap.add_argument("--hasta", required=True, help="YYYY-MM-DD (incl.)")
    ap.add_argument(
        "--codigo-operacion",
        default="1",
        help="Tabla C ARCA: 1 = Retención (recomendado).",
    )
    ap.add_argument(
        "--skip-validacion",
        action="store_true",
        help="No ejecutar --validar-posiciones al final.",
    )
    args, passthrough = ap.parse_known_args(argv)

    root = Path(__file__).resolve().parent
    gen = root / "generar_sicore_v9_retenciones.py"
    out = root / "out" / f"SICORE_V9_RET_GAN_{args.desde}_a_{args.hasta}.TXT"

    cmd = [
        sys.executable,
        str(gen),
        "--desde",
        args.desde,
        "--hasta",
        args.hasta,
        "--out",
        str(out),
        "--layout",
        "estandar132",
        "--codigo-operacion",
        str(args.codigo_operacion),
        *passthrough,
    ]
    r = subprocess.run(cmd)
    if r.returncode != 0:
        return int(r.returncode)

    if not args.skip_validacion:
        v = subprocess.run([sys.executable, str(gen), "--validar-posiciones", str(out)])
        return int(v.returncode)
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
