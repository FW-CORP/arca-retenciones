from __future__ import annotations

import re
import xmlrpc.client
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any


DEC2 = Decimal("0.01")


def d2(x: Any) -> Decimal:
    if x is None or x is False:
        return Decimal("0")
    try:
        return Decimal(str(x))
    except Exception:
        return Decimal("0")


def abs_money(x: Any) -> Decimal:
    return abs(d2(x)).quantize(DEC2, rounding=ROUND_HALF_UP)


def digits_only(s: str) -> str:
    return re.sub(r"\D+", "", s or "")


def ddmmyyyy(iso_date: str) -> str:
    dt = datetime.strptime(iso_date, "%Y-%m-%d").date()
    return dt.strftime("%d/%m/%Y")


def cuit11_from_partner(partner: dict) -> str:
    raw = partner.get("l10n_ar_vat") or partner.get("vat") or ""
    dig = digits_only(str(raw))
    if len(dig) == 11:
        return dig
    if len(dig) > 11:
        return dig[-11:]
    return ""


def is_iibb_tax(tax: dict) -> bool:
    nm = (tax.get("name") or "").upper()
    tt = (tax.get("l10n_ar_tax_type") or "").lower()
    return tt.startswith("iibb") or ("SIRCAR" in nm) or ("IIBB" in nm)


def calc_importe_from_base_rate(base: Decimal, rate_percent: Any) -> Decimal:
    rate = d2(rate_percent)
    return (base * rate / Decimal("100")).quantize(DEC2, rounding=ROUND_HALF_UP)


def odoo_connect(cfg: dict) -> tuple[Any, int]:
    url = cfg["url"].rstrip("/")
    common = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/common", allow_none=True)
    uid = common.authenticate(cfg["db"], cfg["username"], cfg["password"], {})
    if not uid:
        raise SystemExit("Autenticación Odoo fallida (revisá ODOO_CONFIG_MASTER_DEV).")
    return xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/object", allow_none=True), int(uid)


@dataclass(frozen=True)
class RetencionRow:
    payment_id: int
    payment_name: str
    payment_date: str
    partner_name: str
    partner_cuit: str
    tax_name: str
    tax_rate: Decimal
    base: Decimal
    importe: Decimal
    regimen: str
    extra: str


def resolve_project_root(start: Path) -> Path:
    p = start.resolve()
    while p != p.parent:
        if (p / "SICORE" / "run_quincena.py").is_file():
            return p
        p = p.parent
    raise SystemExit("No se encontró la raíz de ARCA-RETENCIONES (falta SICORE/run_quincena.py).")

