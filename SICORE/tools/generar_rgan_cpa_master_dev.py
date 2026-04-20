#!/usr/bin/env python3
"""
Genera `RGAN_CPA.TXT` (ancho fijo) consultando Odoo `master_dev` por XML-RPC.

Solo lectura: no escribe nada en Odoo.

Layout inferido desde `arca-retenciones/RGAN_CPA.TXT` (len 145 por línea):

  0:2    tipo_registro = "06"
  2:12   fecha_1 (10)  = dd/mm/yyyy alineado a derecha (día 1 dígito -> espacio)
  12:16  sucursal (4)  = "0001"
  16:28  nro_orden (12)= correlativo (en el original coincide con otros TXT)
  28:33  espacios (5)
  33:45  importe_total (12) con coma decimal y 3 decimales (ej " 9754137,410")
  45:53  codigo_8 (8)  = "02170781" (configurable)
  53:66  base (13)     con coma decimal y 2 decimales (right aligned)
  66:76  fecha_2 (10)  dd/mm/yyyy alineado a derecha
  76:79  jurisd (3)    = "010" (configurable)
  79:93  retenido (14) con coma decimal y 2 decimales (right aligned)
  93:95  espacios (2)
  95:99  otro (4)      = "0,00" (constante)
  99:109 fecha_3 (10)  dd/mm/yyyy alineado a derecha
  109:122 cuit13 (13)  = "80" + CUIT(11) del proveedor (solo dígitos)
  122:131 espacios (9)
  131:145 ceros (14)   = "00000000000000"

Fuente en Odoo:
- Retenciones en pagos: `account.payment.l10n_ar_withholding_ids` (account.move.line)
- Ganancias: `account.tax.l10n_ar_tax_type == 'earnings'`
"""

from __future__ import annotations

import argparse
import re
import sys
import xmlrpc.client
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Iterable


CURSOR_FILES_ROOT = Path("/media/klap/raid5/cursor_files")
sys.path.insert(0, str(CURSOR_FILES_ROOT))

try:
    from config_nakel import ODOO_CONFIG_MASTER_DEV  # type: ignore
except Exception as e:  # pragma: no cover
    raise SystemExit("No se pudo importar config_nakel.py") from e


DEC2 = Decimal("0.01")
DEC3 = Decimal("0.001")


def _d(x: Any) -> Decimal:
    if x is None or x is False:
        return Decimal("0")
    try:
        return Decimal(str(x))
    except Exception:
        return Decimal("0")


def _abs_money(x: Any, q: Decimal) -> Decimal:
    return abs(_d(x)).quantize(q, rounding=ROUND_HALF_UP)


def _digits_only(s: str) -> str:
    return re.sub(r"\D+", "", s or "")

def _normalize_payment_name_no_separators(name: str, *, max_len: int = 12) -> str:
    """
    Normaliza `account.payment.name` para usar como "nro_orden/nro_comprobante" sin separadores.
    Ej: PGAL1/26-27/0370 -> PGAL126270370 -> recortado a 12 (layout RGAN_CPA usa 12).
    """
    if not name:
        return ""
    out = re.sub(r"[^A-Za-z0-9]+", "", str(name))
    return out[:max_len]

def _fmt_date_ddmmyyyy_right(iso_date: str, width: int = 10) -> str:
    dt = datetime.strptime(iso_date, "%Y-%m-%d").date()
    s = dt.strftime("%d/%m/%Y").lstrip("0")  # 02 -> 2 para parecerse al original
    # si quedó "2/03/2026" (9 chars) lo alineamos a derecha en 10
    return s.rjust(width)


def _fmt_num_coma(x: Decimal, *, decimals: int) -> str:
    # devuelve con coma decimal, sin separador de miles
    q = DEC3 if decimals == 3 else DEC2
    d = x.quantize(q, rounding=ROUND_HALF_UP)
    s = f"{d:.{decimals}f}"
    return s.replace(".", ",")


def _fmt_field_right(value: str, width: int) -> str:
    return (value or "").rjust(width)[:width]


def odoo_connect(cfg: dict) -> tuple[Any, int]:
    url = cfg["url"].rstrip("/")
    common = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/common", allow_none=True)
    uid = common.authenticate(cfg["db"], cfg["username"], cfg["password"], {})
    if not uid:
        raise SystemExit("Autenticación Odoo fallida (ODOO_CONFIG_MASTER_DEV).")
    return xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/object", allow_none=True), int(uid)


def generar(
    desde: str,
    hasta: str,
    *,
    out_path: Path,
    sucursal: str = "0001",
    codigo_8: str = "02170781",
    jurisd_3: str = "010",
) -> Path:
    cfg = ODOO_CONFIG_MASTER_DEV
    models, uid = odoo_connect(cfg)
    db, pwd = cfg["db"], cfg["password"]

    # taxes ganancias
    tax_ids: list[int] = models.execute_kw(
        db,
        uid,
        pwd,
        "account.tax",
        "search",
        [[("l10n_ar_tax_type", "=", "earnings")]],
    )
    if not tax_ids:
        raise SystemExit("No hay impuestos con l10n_ar_tax_type='earnings' en la base.")
    earn_set = set(int(x) for x in tax_ids)

    pay_domain = [
        ("date", ">=", desde),
        ("date", "<=", hasta),
        ("partner_type", "=", "supplier"),
        ("l10n_ar_withholding_ids", "!=", False),
    ]
    pay_ids: list[int] = models.execute_kw(
        db, uid, pwd, "account.payment", "search", [pay_domain], {"order": "date asc, id asc"}
    )
    if not pay_ids:
        raise SystemExit("No se encontraron pagos con retenciones en el rango dado.")

    pays: list[dict] = models.execute_kw(
        db,
        uid,
        pwd,
        "account.payment",
        "read",
        [pay_ids],
        {"fields": ["id", "name", "date", "partner_id", "amount", "l10n_ar_withholding_ids"]},
    )

    partner_ids = sorted({p["partner_id"][0] for p in pays if p.get("partner_id")})
    partners: dict[int, dict] = {}
    if partner_ids:
        pr = models.execute_kw(
            db,
            uid,
            pwd,
            "res.partner",
            "read",
            [partner_ids],
            {"fields": ["id", "vat", "l10n_ar_vat"]},
        )
        partners = {int(r["id"]): r for r in pr}

    line_ids = sorted({lid for p in pays for lid in (p.get("l10n_ar_withholding_ids") or [])})
    lines: list[dict] = models.execute_kw(
        db,
        uid,
        pwd,
        "account.move.line",
        "read",
        [line_ids],
        {"fields": ["id", "payment_id", "tax_line_id", "tax_base_amount", "balance", "credit"]},
    )

    lines_by_payment: dict[int, list[dict]] = {}
    for l in lines:
        if not l.get("tax_line_id"):
            continue
        tid = int(l["tax_line_id"][0])
        if tid not in earn_set:
            continue
        pid = l["payment_id"][0] if l.get("payment_id") else None
        if not pid:
            continue
        lines_by_payment.setdefault(int(pid), []).append(l)

    out_lines: list[str] = []
    for p in pays:
        pid = int(p["id"])
        wlines = lines_by_payment.get(pid) or []
        if not wlines:
            continue

        partner_id = int(p["partner_id"][0]) if p.get("partner_id") else 0
        partner = partners.get(partner_id, {})
        cuit11 = _digits_only(str(partner.get("l10n_ar_vat") or partner.get("vat") or ""))
        if len(cuit11) > 11:
            cuit11 = cuit11[-11:]
        cuit13 = ("80" + cuit11) if len(cuit11) == 11 else "80".ljust(13, "0")

        fecha = _fmt_date_ddmmyyyy_right(p["date"], 10)
        # nro_orden 12: usamos payment.name normalizado (sin / ni -) para alinearnos a SICORE "sin separadores".
        nro_orden = _normalize_payment_name_no_separators(p.get("name") or "", max_len=12).rjust(12, "0") or str(pid).zfill(12)

        imp_total = _abs_money(p.get("amount"), DEC3)
        imp_total_s = _fmt_field_right(_fmt_num_coma(imp_total, decimals=3), 12)

        base_fecha = p["date"]
        fecha2 = _fmt_date_ddmmyyyy_right(base_fecha, 10)
        fecha3 = _fmt_date_ddmmyyyy_right(base_fecha, 10)

        for l in sorted(wlines, key=lambda r: int(r["id"])):
            base = _abs_money(l.get("tax_base_amount"), DEC2)
            reten = _abs_money(l.get("credit") or l.get("balance") or 0, DEC2)
            if reten == Decimal("0.00"):
                reten = _abs_money(l.get("balance"), DEC2)

            base_s = _fmt_field_right(_fmt_num_coma(base, decimals=2), 13)
            reten_s = _fmt_field_right(_fmt_num_coma(reten, decimals=2), 14)

            line = (
                "06"
                + fecha
                + str(sucursal).zfill(4)[:4]
                + nro_orden
                + (" " * 5)
                + imp_total_s
                + str(codigo_8).zfill(8)[:8]
                + base_s
                + fecha2
                + str(jurisd_3).zfill(3)[:3]
                + reten_s
                + (" " * 2)
                + "0,00"
                + fecha3
                + cuit13[:13]
                + (" " * 9)
                + ("0" * 14)
            )
            out_lines.append(line[:145].ljust(145))

    if not out_lines:
        raise SystemExit("No se encontraron retenciones de Ganancias (earnings) en el rango.")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", encoding="utf-8", newline="") as f:
        for ln in out_lines:
            f.write(ln + "\n")

    return out_path


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
        default=Path(__file__).resolve().parent.parent / "out" / "RGAN_CPA.TXT",
    )
    ap.add_argument("--sucursal", default="0001")
    ap.add_argument("--codigo-8", default="02170781")
    ap.add_argument("--jurisd-3", default="010")
    args = ap.parse_args(list(argv) if argv is not None else None)

    out = generar(
        args.desde,
        args.hasta,
        out_path=args.out,
        sucursal=str(args.sucursal),
        codigo_8=str(args.codigo_8),
        jurisd_3=str(args.jurisd_3),
    )
    print(f"OK: {out}")
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())

