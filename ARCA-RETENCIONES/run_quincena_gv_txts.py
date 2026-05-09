#!/usr/bin/env python3
"""
Genera los 3 TXT "legacy" (mismos nombres que el ERP viejo GV) desde Odoo (solo lectura).

Salida (en `ARCA-RETENCIONES/out/`):
- RET-DGR.TXT         (DGR/IIBB ancho fijo)
- RET-DGR-SIRCAR.TXT  (DGR/IIBB CSV tipo SIRCAR)
- RGAN_CPA.TXT        (Ganancias ancho fijo 145)

Ejemplo:
  python3 run_quincena_gv_txts.py --desde 2026-04-01 --hasta 2026-04-15
"""

from __future__ import annotations

import argparse
import subprocess
import sys
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parent

def _default_out_dir() -> Path:
    """
    Para evitar confusiones, por defecto escribimos los 3 TXT en:
      nakel/arca-retenciones/ARCA-RETENCIONES/
    (el directorio donde el equipo suele “mirar”/enviar los archivos).

    Si no se encuentra, caemos a `ARCA-RETENCIONES/out/`.
    """
    for parent in ROOT.parents:
        cand = parent / "arca-retenciones" / "ARCA-RETENCIONES"
        if cand.is_dir():
            return cand
    return ROOT / "out"


def _iso(s: str) -> str:
    datetime.strptime(s, "%Y-%m-%d")
    return s


def _run(cmd: list[str]) -> None:
    p = subprocess.run(cmd, cwd=str(ROOT))
    if p.returncode != 0:
        raise SystemExit(p.returncode)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--desde", type=_iso, required=True, help="YYYY-MM-DD (incl.)")
    ap.add_argument("--hasta", type=_iso, required=True, help="YYYY-MM-DD (incl.)")
    ap.add_argument("--sucursal", default="0001", help="Sucursal (4 dígitos) para layouts fijos")
    ap.add_argument(
        "--header-suffix-code",
        default="240603",
        help="Código de 6 dígitos al final de la leyenda del header de RET-DGR.TXT (observado en el legacy).",
    )
    ap.add_argument("--codigo-regimen", default="001", help="RET-DGR-SIRCAR col 10 (ej. 001)")
    ap.add_argument("--codigo-extra", default="907", help="RET-DGR-SIRCAR col 11 (ej. 907)")
    ap.add_argument(
        "--codigo-8",
        default="2170781",
        help="RGAN_CPA campo código_8 (en legacy observado: 7 dígitos + espacio, ej. 2170781)",
    )
    ap.add_argument("--jurisd-3", default="010", help="RGAN_CPA jurisdicción (3 dígitos)")
    ap.add_argument(
        "--out-dir",
        type=Path,
        default=_default_out_dir(),
        help="Directorio destino de los 3 TXT (por defecto: arca-retenciones/ARCA-RETENCIONES/ si existe).",
    )
    args = ap.parse_args(argv)

    out_dir: Path = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    out_ret_dgr = out_dir / "RET-DGR.TXT"
    out_ret_dgr_sircar = out_dir / "RET-DGR-SIRCAR.TXT"
    out_rgan = out_dir / "RGAN_CPA.TXT"

    _run(
        [
            sys.executable,
            str(ROOT / "SIRCAR" / "tools" / "generar_ret_dgr_ancho_fijo_master_dev.py"),
            "--desde",
            args.desde,
            "--hasta",
            args.hasta,
            "--out",
            str(out_ret_dgr),
            "--sucursal",
            str(args.sucursal),
            "--header-suffix-code",
            str(args.header_suffix_code),
        ]
    )

    _run(
        [
            sys.executable,
            str(ROOT / "SIRCAR" / "tools" / "generar_ret_dgr_master_dev.py"),
            "--desde",
            args.desde,
            "--hasta",
            args.hasta,
            "--out",
            str(out_ret_dgr_sircar),
            "--codigo-regimen",
            str(args.codigo_regimen),
            "--codigo-extra",
            str(args.codigo_extra),
        ]
    )

    _run(
        [
            sys.executable,
            str(ROOT / "SICORE" / "tools" / "generar_rgan_cpa_master_dev.py"),
            "--desde",
            args.desde,
            "--hasta",
            args.hasta,
            "--out",
            str(out_rgan),
            "--sucursal",
            str(args.sucursal),
            "--codigo-8",
            str(args.codigo_8),
            "--jurisd-3",
            str(args.jurisd_3),
        ]
    )

    print("OK:")
    print(f"- {out_ret_dgr}")
    print(f"- {out_ret_dgr_sircar}")
    print(f"- {out_rgan}")
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())

