"""
Bootstrap portable para importar `config_nakel` sin rutas absolutas fijas.

Los scripts añaden la raíz `ARCA-RETENCIONES/` a `sys.path` y luego importan
`prepend_config_nakel_sys_path` desde este módulo.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path


def prepend_config_nakel_sys_path(arca_retenciones_root: Path) -> None:
    """
    Inserta en `sys.path` el directorio que contiene `config_nakel.py`.

    Orden:
    1. `NAKEL_CONFIG_ROOT` (directorio absoluto o relativo al cwd) si apunta a un dir con `config_nakel.py`.
    2. Búsqueda hacia arriba desde `arca_retenciones_root` (incluido) hasta la raíz del filesystem.
    """
    arca = arca_retenciones_root.resolve()
    explicit = os.environ.get("NAKEL_CONFIG_ROOT", "").strip()
    if explicit:
        p = Path(explicit).expanduser()
        if not p.is_absolute():
            p = (Path.cwd() / p).resolve()
        else:
            p = p.resolve()
        if (p / "config_nakel.py").is_file():
            s = str(p)
            if s not in sys.path:
                sys.path.insert(0, s)
            return

    for d in [arca, *arca.parents]:
        if (d / "config_nakel.py").is_file():
            s = str(d)
            if s not in sys.path:
                sys.path.insert(0, s)
            return

    raise SystemExit(
        "No se encontró config_nakel.py.\n"
        "Opciones:\n"
        "  • export NAKEL_CONFIG_ROOT=/ruta/al/directorio/que/contiene/config_nakel.py\n"
        "  • export PYTHONPATH=/esa/misma/ruta:$PYTHONPATH\n"
        "Ver README en ARCA-RETENCIONES (sección Configuración Odoo)."
    )
