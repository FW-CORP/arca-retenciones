#!/usr/bin/env python3
"""
Export SIRCAR (IIBB) desde Odoo `master_dev` (XML-RPC) en modo **solo lectura**.

Este script genera un listado tipo “mayor” con las columnas que aparecen en la captura:
- Fecha
- Cuenta
- Débito
- Crédito
- Etiqueta (detalle breve / referencia)

Filtrado:
- Solo líneas `account.move.line` con `tax_line_id` perteneciente a impuestos IIBB/SIRCAR:
  - `account.tax.l10n_ar_tax_type ilike iibb` OR `name ilike IIBB` OR `name ilike SIRCAR`
- Solo asientos publicados (`parent_state == 'posted'`).

Salida:
- CSV (UTF-8) en `SIRCAR/out/SIRCAR_mayor.csv` por defecto.
"""

from __future__ import annotations

import argparse
import csv
import sys
import xmlrpc.client
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Iterable

_ARCA_SEARCH = Path(__file__).resolve().parent
while _ARCA_SEARCH != _ARCA_SEARCH.parent:
    if (_ARCA_SEARCH / "SICORE" / "run_quincena.py").is_file():
        break
    _ARCA_SEARCH = _ARCA_SEARCH.parent
else:
    raise SystemExit(
        "No se encontró la carpeta del proyecto (debe existir SICORE/run_quincena.py). "
        "Ejecutá los scripts desde el clon …/ARCA-RETENCIONES/ (ver README)."
    )
sys.path.insert(0, str(_ARCA_SEARCH))

try:
    from nakel_import_paths import prepend_config_nakel_sys_path
except Exception as e:  # pragma: no cover
    raise SystemExit("Falta nakel_import_paths.py en la raíz de ARCA-RETENCIONES.") from e

prepend_config_nakel_sys_path(_ARCA_SEARCH)

try:
    from config_nakel import ODOO_CONFIG_MASTER_DEV  # type: ignore
except Exception as e:  # pragma: no cover
    raise SystemExit(
        "No se pudo importar config_nakel (¿NAKEL_CONFIG_ROOT o PYTHONPATH?). Ver README de ARCA-RETENCIONES."
    ) from e

DEC2 = Decimal("0.01")


def _d2(x: Any) -> Decimal:
    if x is None or x is False:
        return Decimal("0")
    try:
        return Decimal(str(x))
    except Exception:
        return Decimal("0")


def _fmt_money_cell(x: Any) -> str:
    q = _d2(x).quantize(DEC2, rounding=ROUND_HALF_UP)
    s = f"{q:.2f}"
    return "0" if s in ("0.00", "-0.00") else s.rstrip("0").rstrip(".")


def _ddmmyyyy(iso_date: str) -> str:
    dt = datetime.strptime(iso_date, "%Y-%m-%d").date()
    return dt.strftime("%d/%m/%Y")


def odoo_connect(cfg: dict) -> tuple[Any, int]:
    url = cfg["url"].rstrip("/")
    common = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/common", allow_none=True)
    uid = common.authenticate(cfg["db"], cfg["username"], cfg["password"], {})
    if not uid:
        raise SystemExit("Autenticación Odoo fallida (ODOO_CONFIG_MASTER_DEV).")
    return xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/object", allow_none=True), int(uid)


def _iibb_tax_ids(models: Any, db: str, uid: int, pwd: str) -> list[int]:
    dom: list[Any] = [
        "|",
        "|",
        ("l10n_ar_tax_type", "ilike", "iibb"),
        ("name", "ilike", "IIBB"),
        ("name", "ilike", "SIRCAR"),
    ]
    ids: list[int] = models.execute_kw(db, uid, pwd, "account.tax", "search", [dom])
    return [int(x) for x in ids]


def generar(desde: str, hasta: str, *, out_csv: Path) -> Path:
    cfg = ODOO_CONFIG_MASTER_DEV
    models, uid = odoo_connect(cfg)
    db, pwd = cfg["db"], cfg["password"]

    tax_ids = _iibb_tax_ids(models, db, uid, pwd)
    if not tax_ids:
        raise SystemExit("No se encontraron impuestos IIBB/SIRCAR en Odoo (account.tax).")

    dom = [
        ("date", ">=", desde),
        ("date", "<=", hasta),
        ("tax_line_id", "in", tax_ids),
        ("parent_state", "=", "posted"),
    ]
    aml_ids: list[int] = models.execute_kw(db, uid, pwd, "account.move.line", "search", [dom], {"order": "date asc, id asc"})
    if not aml_ids:
        raise SystemExit("No se encontraron líneas IIBB/SIRCAR en el rango dado.")

    rows: list[dict] = models.execute_kw(
        db,
        uid,
        pwd,
        "account.move.line",
        "read",
        [aml_ids],
        {
            "fields": [
                "id",
                "date",
                "account_id",
                "debit",
                "credit",
                "name",
                "ref",
                "move_name",
                "payment_id",
            ]
        },
    )

    out_csv.parent.mkdir(parents=True, exist_ok=True)
    with out_csv.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Fecha", "Cuenta", "Debito", "Credito", "Etiqueta"])
        for r in rows:
            fecha = _ddmmyyyy(str(r["date"]))
            cuenta = str(r["account_id"][1]) if r.get("account_id") else ""
            deb = _fmt_money_cell(r.get("debit"))
            cre = _fmt_money_cell(r.get("credit"))
            etiqueta = str(r.get("name") or r.get("ref") or r.get("move_name") or "")
            w.writerow([fecha, cuenta, deb, cre, etiqueta])

    return out_csv


def _iso(s: str) -> str:
    datetime.strptime(s, "%Y-%m-%d")
    return s


def _default_range_mes_actual() -> tuple[str, str]:
    hoy = date.today()
    desde = hoy.replace(day=1)
    if desde.month == 12:
        first_next = desde.replace(year=desde.year + 1, month=1, day=1)
    else:
        first_next = desde.replace(month=desde.month + 1, day=1)
    from datetime import timedelta

    hasta = first_next - timedelta(days=1)
    return desde.isoformat(), hasta.isoformat()


def main(argv: Iterable[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    d_desde, d_hasta = _default_range_mes_actual()
    ap.add_argument("--desde", type=_iso, default=d_desde)
    ap.add_argument("--hasta", type=_iso, default=d_hasta)
    ap.add_argument(
        "--out",
        type=Path,
        default=Path(__file__).resolve().parent.parent / "out" / "SIRCAR_mayor.csv",
    )
    args = ap.parse_args(list(argv) if argv is not None else None)

    out = generar(args.desde, args.hasta, out_csv=args.out)
    print(f"OK: {out}")
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())

