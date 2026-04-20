#!/usr/bin/env python3
"""
Genera `RET-DGR.TXT` (ancho fijo) consultando Odoo `master_dev` por XML-RPC.

Solo lectura: no escribe nada en Odoo.

Layout inferido desde el TXT original del contador (Nakel).

HEADER (1 línea, len 175):
  - 0:7    periodo MM/YYYY
  - 7:21   total (14 dígitos) -> total retenido en centavos (default)
  - 21:76  leyenda en mayúsculas (55 chars) -> "SON PESOS ..." + sufijo configurable (6 chars)
  - 76:89  CUIT agente (formato 30-XXXXXXXX-X)
  - 89:119 razón social (30)
  - 119:149 domicilio (30)
  - 149:164 localidad (15)
  - 164:175 relleno (11 ceros)

DETALLE (N líneas, len 184):
  - 0:10   nro_orden (10) -> payment.id zfill(10)
  - 10:20  fecha dd/mm/yyyy -> payment.date
  - 20:24  sucursal (4) -> "0001"
  - 24:32  nro_interno (8) -> correlativo de detalle (zfill(8))
  - 32     espacio
  - 33:69  3 importes * 12 dígitos (centavos, sin separadores):
           (1) importe_total_pago
           (2) base_retencion
           (3) importe_retenido
  - 69:94  "000-000000-" + CUIT proveedor en formato "0DD-XXXXXXXX-X" (len 25)
  - 94:124 nombre proveedor (30)
  - 124:154 domicilio proveedor (30)
  - 154:174 localidad proveedor (20)
  - 174:182 CP proveedor (8)
  - 182:184 código provincia AFIP (2) (mapeo fijo por nombre de provincia)
"""

from __future__ import annotations

import argparse
import re
import sys
import xmlrpc.client
from dataclasses import dataclass
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


AFIP_PROV_CODE_BY_STATE_NAME: dict[str, str] = {
    "Buenos Aires": "02",
    "Ciudad Autónoma de Buenos Aires": "00",
    "Capital Federal": "00",
    "Catamarca": "03",
    "Córdoba": "04",
    "Cordoba": "04",
    "Corrientes": "05",
    "Chaco": "06",
    "Chubut": "07",
    "Entre Ríos": "08",
    "Entre Rios": "08",
    "Formosa": "09",
    "Jujuy": "10",
    "La Pampa": "11",
    "La Rioja": "12",
    "Mendoza": "13",
    "Misiones": "14",
    "Neuquén": "15",
    "Neuquen": "15",
    "Río Negro": "16",
    "Rio Negro": "16",
    "Salta": "17",
    "San Juan": "18",
    "San Luis": "19",
    "Santa Cruz": "20",
    "Santa Fe": "21",
    "Santiago del Estero": "22",
    "Tierra del Fuego": "23",
    "Tucumán": "24",
    "Tucuman": "24",
}

def _norm_state_name(name: str) -> str:
    n = (name or "").strip()
    # Odoo suele tener "Chubut (AR)".
    if n.endswith("(AR)"):
        n = n[:-4].strip()
    # Normalizar variantes comunes sin acentos en el diccionario
    return n


def _d2(x: Any) -> Decimal:
    if x is None or x is False:
        return Decimal("0")
    try:
        return Decimal(str(x))
    except Exception:
        return Decimal("0")


def _abs_money(x: Any) -> Decimal:
    return abs(_d2(x)).quantize(DEC2, rounding=ROUND_HALF_UP)


def _digits_only(s: str) -> str:
    return re.sub(r"\D+", "", s or "")

def _normalize_payment_name_no_separators(name: str) -> str:
    """
    Normaliza `account.payment.name` quitando separadores (/,-,espacios,etc).
    Ej: PGAL1/26-27/0370 -> PGAL126270370
    """
    if not name:
        return ""
    return re.sub(r"[^A-Za-z0-9]+", "", str(name))

def _nro_orden_10_from_payment_name(name: str) -> str:
    """
    En este layout el campo es de 10 chars.
    Usamos el `payment.name` normalizado y:
    - si supera 10, tomamos los últimos 10 (prioriza el sufijo secuencial)
    - si es menor, rellenamos con ceros a la izquierda
    """
    clean = _normalize_payment_name_no_separators(name)
    if not clean:
        return ""
    if len(clean) > 10:
        clean = clean[-10:]
    return clean.rjust(10, "0")


def _ddmmyyyy(iso_date: str) -> str:
    dt = datetime.strptime(iso_date, "%Y-%m-%d").date()
    return dt.strftime("%d/%m/%Y")


def _mmYYYY(iso_date: str) -> str:
    dt = datetime.strptime(iso_date, "%Y-%m-%d").date()
    return dt.strftime("%m/%Y")


def _cents12(x: Decimal) -> str:
    cents = int((x * 100).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    return str(cents).zfill(12)


def _pad(s: str, width: int) -> str:
    s2 = (s or "")[:width]
    return s2.ljust(width)


def _fmt_cuit_std(cuit11: str) -> str:
    d = _digits_only(cuit11)
    if len(d) != 11:
        return ""
    return f"{d[:2]}-{d[2:10]}-{d[10:]}"


def _fmt_cuit_pref_0dd(cuit11: str) -> str:
    d = _digits_only(cuit11)
    if len(d) != 11:
        return ""
    return f"0{d[:2]}-{d[2:10]}-{d[10:]}"


def _partner_cuit11(partner: dict) -> str:
    raw = partner.get("l10n_ar_vat") or partner.get("vat") or ""
    dig = _digits_only(str(raw))
    if len(dig) == 11:
        return dig
    if len(dig) > 11:
        return dig[-11:]
    return ""


def _prov_afip_code(partner: dict) -> str:
    st = partner.get("state_id")
    if st and isinstance(st, list) and len(st) >= 2:
        name = _norm_state_name(str(st[1]))
        return AFIP_PROV_CODE_BY_STATE_NAME.get(name, "00")
    return "00"


def _partner_localidad(partner: dict) -> str:
    city = (partner.get("city") or "").strip()
    if city:
        return city
    st = partner.get("state_id")
    if st and isinstance(st, list) and len(st) >= 2:
        return _norm_state_name(str(st[1]))
    return ""


def _partner_domicilio(partner: dict) -> str:
    street = (partner.get("street") or "").strip()
    street2 = (partner.get("street2") or "").strip()
    if street2:
        return f"{street} {street2}".strip()
    return street


def odoo_connect(cfg: dict) -> tuple[Any, int]:
    url = cfg["url"].rstrip("/")
    common = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/common", allow_none=True)
    uid = common.authenticate(cfg["db"], cfg["username"], cfg["password"], {})
    if not uid:
        raise SystemExit("Autenticación Odoo fallida (revisá ODOO_CONFIG_MASTER_DEV).")
    return xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/object", allow_none=True), int(uid)


@dataclass(frozen=True)
class Detalle:
    nro_orden_10: str
    fecha_10: str
    suc_4: str
    nro_interno_8: str
    imp_total_12: str
    base_12: str
    retenido_12: str
    cuit_block_25: str
    nombre_30: str
    domicilio_30: str
    localidad_20: str
    cp_8: str
    prov_2: str

    def render(self) -> str:
        return (
            f"{self.nro_orden_10}"
            f"{self.fecha_10}"
            f"{self.suc_4}"
            f"{self.nro_interno_8}"
            f" "
            f"{self.imp_total_12}{self.base_12}{self.retenido_12}"
            f"{self.cuit_block_25}"
            f"{self.nombre_30}"
            f"{self.domicilio_30}"
            f"{self.localidad_20}"
            f"{self.cp_8}"
            f"{self.prov_2}"
        )


def _is_iibb_sircar_tax(tax: dict) -> bool:
    nm = (tax.get("name") or "").upper()
    tt = (tax.get("l10n_ar_tax_type") or "").lower()
    return tt.startswith("iibb") or ("SIRCAR" in nm) or ("IIBB" in nm)


def generar(
    desde: str,
    hasta: str,
    *,
    out_path: Path,
    sucursal: str = "0001",
    header_suffix_code: str = "240603",
) -> Path:
    cfg = ODOO_CONFIG_MASTER_DEV
    models, uid = odoo_connect(cfg)
    db, pwd = cfg["db"], cfg["password"]

    # Empresa (header)
    company: dict = {}
    comps = models.execute_kw(
        db,
        uid,
        pwd,
        "res.company",
        "search_read",
        [[]],
        {"fields": ["id", "name", "vat", "street", "street2", "city", "zip", "state_id"], "order": "id asc"},
    )
    # Elegir la compañía "real": la primera con VAT/CUIT válido; si no, la primera.
    for c in comps:
        vat = _digits_only(str(c.get("vat") or ""))
        if len(vat) == 11:
            company = c
            break
    if not company and comps:
        company = comps[0]

    # Pagos con retenciones
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

    # Partners
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
            {"fields": ["id", "name", "vat", "l10n_ar_vat", "street", "street2", "city", "zip", "state_id"]},
        )
        partners = {int(r["id"]): r for r in pr}

    # Withholding lines
    line_ids = sorted({lid for p in pays for lid in (p.get("l10n_ar_withholding_ids") or [])})
    lines: list[dict] = models.execute_kw(
        db,
        uid,
        pwd,
        "account.move.line",
        "read",
        [line_ids],
        {"fields": ["id", "payment_id", "tax_base_amount", "balance", "credit", "tax_line_id"]},
    )

    tax_ids = sorted({l["tax_line_id"][0] for l in lines if l.get("tax_line_id")})
    taxes: dict[int, dict] = {}
    if tax_ids:
        tr = models.execute_kw(
            db,
            uid,
            pwd,
            "account.tax",
            "read",
            [tax_ids],
            {"fields": ["id", "name", "l10n_ar_tax_type"]},
        )
        taxes = {int(r["id"]): r for r in tr}

    lines_by_payment: dict[int, list[dict]] = {}
    for l in lines:
        pid = l["payment_id"][0] if l.get("payment_id") else None
        if not pid:
            continue
        tax = taxes.get(l["tax_line_id"][0]) if l.get("tax_line_id") else None
        if not tax or not _is_iibb_sircar_tax(tax):
            continue
        lines_by_payment.setdefault(int(pid), []).append(l)

    detalles: list[Detalle] = []
    nro_interno = 0
    total_retenido = Decimal("0.00")
    periodo = _mmYYYY(pays[0]["date"])

    for p in pays:
        pid = int(p["id"])
        wlines = lines_by_payment.get(pid) or []
        if not wlines:
            continue

        partner_id = int(p["partner_id"][0]) if p.get("partner_id") else 0
        partner = partners.get(partner_id, {})
        cuit11 = _partner_cuit11(partner)
        cuit_block = (
            f"000-000000-{_fmt_cuit_pref_0dd(cuit11)}"
            if cuit11
            else "000-000000-000-00000000-0"
        )

        nombre = _pad((partner.get("name") or "").upper(), 30)
        domicilio = _pad(_partner_domicilio(partner), 30)
        localidad = _pad(_partner_localidad(partner), 20)
        cp = _pad((partner.get("zip") or "").strip(), 8)
        prov = _prov_afip_code(partner)

        imp_total = _abs_money(p.get("amount"))
        for l in sorted(wlines, key=lambda r: int(r["id"])):
            base = _abs_money(l.get("tax_base_amount"))
            reten = _abs_money(l.get("credit") or l.get("balance") or 0)
            if reten == Decimal("0.00"):
                reten = _abs_money(l.get("balance"))

            nro_interno += 1
            total_retenido += reten

            detalles.append(
                Detalle(
                    nro_orden_10=_nro_orden_10_from_payment_name(p.get("name") or "") or str(pid).zfill(10),
                    fecha_10=_ddmmyyyy(p["date"]),
                    suc_4=sucursal,
                    nro_interno_8=str(nro_interno).zfill(8),
                    imp_total_12=_cents12(imp_total),
                    base_12=_cents12(base),
                    retenido_12=_cents12(reten),
                    cuit_block_25=cuit_block,
                    nombre_30=nombre,
                    domicilio_30=domicilio,
                    localidad_20=localidad,
                    cp_8=cp,
                    prov_2=prov,
                )
            )

    if not detalles:
        raise SystemExit("No se encontraron retenciones IIBB/SIRCAR en el rango dado.")

    total14 = str(int((total_retenido * 100).quantize(Decimal("1"), rounding=ROUND_HALF_UP))).zfill(14)
    cuit_agente = _fmt_cuit_std(_digits_only(company.get("vat") or ""))
    if not cuit_agente:
        cuit_agente = "00-00000000-0"

    leyenda = f"SON PESOS {total_retenido:,.2f}"
    leyenda = leyenda.replace(",", "X").replace(".", ",").replace("X", ".")
    leyenda = (leyenda.upper() + header_suffix_code)[:55].ljust(55)

    rs = _pad((company.get("name") or "").strip(), 30)
    dom = _pad(_partner_domicilio(company), 30)
    loc = _pad(_partner_localidad(company), 15)
    zeros11 = "0" * 11
    header = f"{periodo}{total14}{leyenda}{cuit_agente}{rs}{dom}{loc}{zeros11}"
    header = header[:175].ljust(175)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", encoding="utf-8", newline="") as f:
        f.write(header + "\n")
        for d in detalles:
            f.write(d.render()[:184].ljust(184) + "\n")

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
        default=Path(__file__).resolve().parent.parent / "out" / "RET-DGR.TXT",
    )
    ap.add_argument("--sucursal", default="0001")
    ap.add_argument(
        "--header-suffix-code",
        default="240603",
        help="Código de 6 dígitos al final de la leyenda del header (observado en el original).",
    )
    args = ap.parse_args(list(argv) if argv is not None else None)

    out = generar(
        args.desde,
        args.hasta,
        out_path=args.out,
        sucursal=str(args.sucursal).zfill(4)[:4],
        header_suffix_code=str(args.header_suffix_code).strip()[:6].ljust(6, "0"),
    )
    print(f"OK: {out}")
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())

