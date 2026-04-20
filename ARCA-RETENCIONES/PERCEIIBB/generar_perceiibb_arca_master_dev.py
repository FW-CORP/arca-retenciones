#!/usr/bin/env python3
"""
Genera TXT SIRCAR **Percepciones IIBB** de ancho fijo **163** caracteres (misma grilla posicional
que retenciones, con diferencias de negocio indicadas abajo).

Referencia layout: tabla percepciones (captura del estudio).

Fuente de datos (solo lectura):
- `account.move` (facturas / NC / ND de **cliente**), `invoice_date` en el rango
- `account.move.line` con `tax_line_id` en impuestos de **percepción** IIBB (`account.tax`)

Diferencias frente a retenciones (`SIRCAR/generar_sircar_163_master_dev.py`):
- **Campo 1**: `2` = percepción (no `1` retención).
- **Campo 6 (fecha)**: fecha de la **factura de venta** (`invoice_date`), no fecha de pago.
- **Campo 4 (mes/año)** y **cuota**: según `invoice_date` de cada comprobante.
- **Campo 8 (tipo comprobante)**: `F` factura, `D` nota de débito, `C` nota de crédito (NC: monto percibido informado en **positivo**; SIRCAR resta al procesar `C`).
- **Sujeto**: **cliente** (`partner_id` de la factura), no proveedor del pago.

Encoding: ASCII; fin de línea CRLF; numéricos con ceros a la izquierda; razón social 40 con espacios a la derecha.
"""

from __future__ import annotations

import argparse
import re
import sys
import unicodedata
import xmlrpc.client
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Iterable


def _arca_retenciones_root() -> Path:
    """Raíz `ARCA-RETENCIONES/` (nakel_import_paths + SICORE/run_quincena.py)."""
    p = Path(__file__).resolve().parent
    for cur in [p, *p.parents]:
        if (cur / "nakel_import_paths.py").is_file() and (cur / "SICORE" / "run_quincena.py").is_file():
            return cur
        nested = cur / "ARCA-RETENCIONES"
        if (nested / "nakel_import_paths.py").is_file() and (nested / "SICORE" / "run_quincena.py").is_file():
            return nested
    raise SystemExit(
        "No se encontró ARCA-RETENCIONES (nakel_import_paths.py + SICORE/run_quincena.py) subiendo desde PERCEIIBB/. "
        "Ejecutá desde el clon bajo …/nakel_scripts/ o …/arca-retenciones/ (ver README)."
    )


_ARCA = _arca_retenciones_root()
sys.path.insert(0, str(_ARCA))

try:
    from nakel_import_paths import prepend_config_nakel_sys_path
except Exception as e:  # pragma: no cover
    raise SystemExit("Falta nakel_import_paths.py en ARCA-RETENCIONES.") from e

prepend_config_nakel_sys_path(_ARCA)

try:
    from config_nakel import ODOO_CONFIG_MASTER_DEV  # type: ignore
except Exception as e:  # pragma: no cover
    raise SystemExit(
        "No se pudo importar config_nakel (¿NAKEL_CONFIG_ROOT o PYTHONPATH?). Ver README de ARCA-RETENCIONES."
    ) from e

DEC3 = Decimal("0.001")
LREG = 163


def odoo_connect(cfg: dict) -> tuple[Any, int]:
    url = cfg["url"].rstrip("/")
    common = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/common", allow_none=True)
    uid = common.authenticate(cfg["db"], cfg["username"], cfg["password"], {})
    if not uid:
        raise SystemExit("Autenticación Odoo fallida (ODOO_CONFIG_MASTER_DEV).")
    return xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/object", allow_none=True), int(uid)


def _digits_only(s: str) -> str:
    return re.sub(r"\D+", "", s or "")


def _ddmmyyyy(iso_date: str) -> str:
    return datetime.strptime(iso_date, "%Y-%m-%d").date().strftime("%d/%m/%Y")


def _yyyymm_from_iso(iso_date: str) -> str:
    return datetime.strptime(iso_date, "%Y-%m-%d").date().strftime("%Y%m")


def _fmt_centavos(value: Any, width: int) -> str:
    d = Decimal(str(value or 0))
    cents = (abs(d) * 100).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
    n = int(cents) if cents is not None else 0
    return str(n).zfill(width)[-width:]


def _fmt_alicuota_5(value_percent: Any) -> str:
    d = Decimal(str(value_percent or 0)).quantize(DEC3, rounding=ROUND_HALF_UP)
    n = int((d * 1000).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    if n < 0:
        n = -n
    return str(n).zfill(5)[-5:]


def _ascii_upper_pad_right(s: str, width: int) -> str:
    s = (s or "").strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if ord(ch) < 128)
    s = s.upper()
    if len(s) >= width:
        return s[:width]
    return s + (" " * (width - len(s)))


def _fiscal_pv_num_from_raw(raw: str) -> tuple[int, int] | None:
    s = (raw or "").strip()
    if not s:
        return None
    m = re.match(r"^\s*(\d{1,9})\s*[-/]\s*(\d{1,12})\s*$", s)
    if not m:
        return None
    pv = max(0, min(int(m.group(1)), 9999))
    num = max(0, min(int(m.group(2)), 999999999999))
    return pv, num


def _fiscal_pv_num_sircar(raw: str) -> tuple[int, int]:
    parsed = _fiscal_pv_num_from_raw(raw)
    if parsed:
        return parsed
    d = _digits_only(raw or "")
    if not d:
        return 0, 0
    if len(d) >= 16:
        chunk = d[-16:]
        pv = max(0, min(int(chunk[:4]), 9999))
        rest = chunk[4:]
        num = max(0, min(int(rest) if rest else 0, 999999999999))
        return pv, num
    if len(d) >= 10:
        pv = max(0, min(int(d[:4]), 9999))
        rest = d[4:]
        num = max(0, min(int(rest) if rest else 0, 999999999999))
        return pv, num
    num = max(0, min(int(d), 999999999999))
    return 0, num


def _nro_comp_12(raw: str) -> str:
    _, num = _fiscal_pv_num_sircar(raw)
    return f"{num:012d}"[-12:]


def _pv_num_16(raw: str) -> str:
    pv, num = _fiscal_pv_num_sircar(raw)
    return f"{pv:04d}{num:012d}"


def _tax_ids_percepcion_iibb(models: Any, db: str, uid: int, pwd: str) -> set[int]:
    """Impuestos de percepción IIBB (`l10n_ar_tax_type` suele contener `perception`)."""
    dom_primary: list[Any] = [("l10n_ar_tax_type", "ilike", "perception")]
    ids = models.execute_kw(db, uid, pwd, "account.tax", "search", [dom_primary])
    out = {int(x) for x in ids}
    if out:
        return out
    dom_fallback: list[Any] = [
        "&",
        ("name", "ilike", "perc"),
        "|",
        ("name", "ilike", "IIBB"),
        ("name", "ilike", "SIRCAR"),
    ]
    ids2 = models.execute_kw(db, uid, pwd, "account.tax", "search", [dom_fallback])
    return {int(x) for x in ids2}


def _juris_9xx_desde_impuesto(tax: dict) -> str | None:
    blob = " ".join(str(tax.get(k) or "") for k in ("name", "l10n_ar_code"))
    found = re.findall(r"(?<!\d)(9\d{2})(?!\d)", blob)
    if not found:
        return None
    for pref in ("907", "920", "901", "902", "903", "904", "905", "906", "908", "909"):
        if pref in found:
            return pref
    return found[0]


def _letra_desde_latam_doc_type(dt: dict | None) -> str | None:
    if not dt:
        return None
    name = str(dt.get("name") or "")
    code = str(dt.get("code") or "")
    name_u = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode().upper()
    code_u = code.upper()
    for blob in (name_u, code_u):
        for letter, needles in (
            ("A", ("FACTURA A", "NOTA DE CREDITO A", "NOTA DE DEBITO A", " LETRA A", "-A", "/A")),
            ("B", ("FACTURA B", "NOTA DE CREDITO B", "NOTA DE DEBITO B", " LETRA B", "-B", "/B")),
            ("C", ("FACTURA C", "NOTA DE CREDITO C", "NOTA DE DEBITO C", " LETRA C", "-C", "/C")),
        ):
            if any(n in blob for n in needles):
                return letter
    m = re.search(r"\b([ABC])\b", name_u)
    if m:
        return m.group(1)
    return None


def _letra_desde_factura(inv: dict, doc_type: dict | None) -> str:
    letra = _letra_desde_latam_doc_type(doc_type)
    if letra:
        return letra
    name_hint = str(inv.get("name") or "")
    for cand in ("FA-A", "FC-A", "NC-A", "ND-A"):
        if cand in name_hint:
            return "A"
    for cand in ("FA-B", "FC-B", "NC-B", "ND-B"):
        if cand in name_hint:
            return "B"
    for cand in ("FA-C", "FC-C", "NC-C", "ND-C"):
        if cand in name_hint:
            return "C"
    return " "


def _tipo_comprobante_8(move_type: str) -> str:
    if move_type == "out_refund":
        return "C"
    if move_type == "out_debit":
        return "D"
    return "F"


def generar(
    desde: str,
    hasta: str,
    *,
    out_path: Path,
    cuit_agente_11: str,
    jurisdiccion_3: str,
    jurisdiccion_sujeto_3: str | None = None,
    cuota_1: str | None = None,
) -> Path:
    cfg = ODOO_CONFIG_MASTER_DEV
    models, uid = odoo_connect(cfg)
    db, pwd = cfg["db"], cfg["password"]

    tax_perc = _tax_ids_percepcion_iibb(models, db, uid, pwd)
    if not tax_perc:
        raise SystemExit(
            "No hay impuestos de percepción IIBB detectados en account.tax "
            "(l10n_ar_tax_type con 'perception' o nombre tipo PERC IIBB). Revisar maestro de impuestos."
        )

    move_domain = [
        ("move_type", "in", ["out_invoice", "out_refund", "out_debit"]),
        ("state", "=", "posted"),
        ("invoice_date", ">=", desde),
        ("invoice_date", "<=", hasta),
    ]
    move_ids: list[int] = models.execute_kw(
        db, uid, pwd, "account.move", "search", [move_domain], {"order": "invoice_date asc, id asc"}
    )
    if not move_ids:
        raise SystemExit("No hay facturas de cliente publicadas en el rango de fechas.")

    line_domain = [
        ("move_id", "in", move_ids),
        ("tax_line_id", "in", sorted(tax_perc)),
    ]
    line_ids: list[int] = models.execute_kw(
        db, uid, pwd, "account.move.line", "search", [line_domain], {"order": "move_id asc, id asc"}
    )
    if not line_ids:
        raise SystemExit(
            "Hay facturas en el rango pero ninguna línea de impuesto de percepción IIBB "
            "(tax_line_id) asociada a esos impuestos."
        )

    lines: list[dict] = models.execute_kw(
        db,
        uid,
        pwd,
        "account.move.line",
        "read",
        [line_ids],
        {"fields": ["id", "move_id", "tax_line_id", "tax_base_amount", "balance", "credit", "debit"]},
    )

    inv_ids = sorted({int(l["move_id"][0]) for l in lines if l.get("move_id")})
    invs: list[dict] = models.execute_kw(
        db,
        uid,
        pwd,
        "account.move",
        "read",
        [inv_ids],
        {
            "fields": [
                "id",
                "name",
                "move_type",
                "invoice_date",
                "partner_id",
                "l10n_latam_document_number",
                "l10n_latam_document_type_id",
            ]
        },
    )
    inv_map = {int(r["id"]): r for r in invs}

    partner_ids = sorted({int(r["partner_id"][0]) for r in invs if r.get("partner_id")})
    partners: dict[int, dict] = {}
    if partner_ids:
        pr = models.execute_kw(
            db,
            uid,
            pwd,
            "res.partner",
            "read",
            [partner_ids],
            {"fields": ["id", "name", "vat", "l10n_ar_vat"]},
        )
        partners = {int(r["id"]): r for r in pr}

    tax_ids = sorted({int(l["tax_line_id"][0]) for l in lines if l.get("tax_line_id")})
    taxes: dict[int, dict] = {}
    if tax_ids:
        tr = models.execute_kw(
            db,
            uid,
            pwd,
            "account.tax",
            "read",
            [tax_ids],
            {"fields": ["id", "name", "amount", "l10n_ar_tax_type", "l10n_ar_code"]},
        )
        taxes = {int(r["id"]): r for r in tr}

    doc_type_ids = sorted(
        {
            int(inv["l10n_latam_document_type_id"][0])
            for inv in inv_map.values()
            if inv.get("l10n_latam_document_type_id")
        }
    )
    doc_types: dict[int, dict] = {}
    if doc_type_ids:
        dr = models.execute_kw(
            db,
            uid,
            pwd,
            "l10n_latam.document.type",
            "read",
            [doc_type_ids],
            {"fields": ["id", "name", "code"]},
        )
        doc_types = {int(r["id"]): r for r in dr}

    juris = str(jurisdiccion_3).zfill(3)[:3]

    out_lines: list[str] = []
    for l in lines:
        mid = l.get("move_id")
        tid = l.get("tax_line_id")
        if not mid or not tid:
            continue
        inv = inv_map.get(int(mid[0]))
        if not inv:
            continue
        tid_i = int(tid[0])
        tax = taxes.get(tid_i) or {}

        inv_date = str(inv.get("invoice_date") or "")[:10]
        if not inv_date:
            raise SystemExit(f"Factura id={inv['id']} sin invoice_date.")
        fecha_fac = _ddmmyyyy(inv_date)
        mes_anio = _yyyymm_from_iso(inv_date)
        cuota = (
            cuota_1
            if cuota_1
            else ("1" if int(inv_date[8:10]) <= 15 else "2")
        )

        raw_nro = str(inv.get("l10n_latam_document_number") or inv.get("name") or "")
        campo7_12 = _nro_comp_12(raw_nro)
        campo10_16 = _pv_num_16(raw_nro)

        tipo_8 = _tipo_comprobante_8(str(inv.get("move_type") or ""))
        dt_id = inv.get("l10n_latam_document_type_id")
        doc_row = doc_types.get(int(dt_id[0])) if dt_id else None
        letra = _letra_desde_factura(inv, doc_row)

        base_15 = _fmt_centavos(abs(Decimal(str(l.get("tax_base_amount") or 0))), 15)
        # Monto percibido siempre positivo en archivo (NC con C: SIRCAR resta al importar)
        bal = Decimal(str(l.get("balance") or 0))
        if bal == 0:
            bal = Decimal(str(l.get("credit") or 0)) - Decimal(str(l.get("debit") or 0))
        monto_15 = _fmt_centavos(abs(bal), 15)
        alic_5 = _fmt_alicuota_5(tax.get("amount"))

        partner_id = int(inv["partner_id"][0]) if inv.get("partner_id") else 0
        partner = partners.get(partner_id, {})
        cuit_suj = _digits_only(str(partner.get("l10n_ar_vat") or partner.get("vat") or "")).rjust(11, "0")[-11:]
        razon_40 = _ascii_upper_pad_right(str(partner.get("name") or ""), 40)

        subj_tax = _juris_9xx_desde_impuesto(tax)
        if jurisdiccion_sujeto_3 is not None:
            juris_suj = str(jurisdiccion_sujeto_3).zfill(3)[:3]
        elif subj_tax:
            juris_suj = subj_tax
        else:
            juris_suj = juris

        tipo_reg = "2"
        tipo_doc_1 = "1"
        otros_11 = "0" * 11

        out = (
            tipo_reg
            + juris
            + _digits_only(cuit_agente_11).rjust(11, "0")[-11:]
            + mes_anio
            + str(cuota)[:1]
            + fecha_fac
            + campo7_12
            + tipo_8
            + letra
            + campo10_16
            + base_15
            + alic_5
            + monto_15
            + tipo_doc_1
            + cuit_suj
            + juris_suj
            + razon_40
            + otros_11
        )
        if len(out) != LREG:
            raise SystemExit(f"Línea {len(out)} != {LREG} (move={inv['id']}, aml={l['id']})")
        out_lines.append(out)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", encoding="ascii", errors="replace", newline="") as f:
        for ln in out_lines:
            f.write(ln + "\r\n")
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
    ap.add_argument("--cuit-agente", required=True, help="CUIT agente de percepción (11 dígitos, sin guiones).")
    ap.add_argument(
        "--jurisdiccion",
        default="907",
        help="Jurisdicción agente / donde se tributa la venta (3 dígitos). Default 907.",
    )
    ap.add_argument(
        "--jurisdiccion-sujeto",
        default=None,
        help="Jurisdicción del cliente (3 dígitos). Si se omite: 9xx desde impuesto; si no, misma que agente.",
    )
    ap.add_argument("--cuota", default=None, help="1 o 2. Si se omite, se calcula por día de invoice_date (<=15 => 1).")
    ap.add_argument(
        "--out",
        type=Path,
        default=Path(__file__).resolve().parent / "out" / "PERCEIIBB_ARCA.TXT",
    )
    args = ap.parse_args(list(argv) if argv is not None else None)

    out = generar(
        args.desde,
        args.hasta,
        out_path=args.out,
        cuit_agente_11=str(args.cuit_agente),
        jurisdiccion_3=str(args.jurisdiccion),
        jurisdiccion_sujeto_3=str(args.jurisdiccion_sujeto) if args.jurisdiccion_sujeto is not None else None,
        cuota_1=str(args.cuota) if args.cuota is not None else None,
    )
    print(f"OK: {out}")
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
