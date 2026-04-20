#!/usr/bin/env python3
"""
Genera TXT SIRCAR (IIBB) de ancho fijo **163** caracteres según layout canónico del estudio.

Referencia layout: `assets/image-df88190b-8204-479b-840d-109d33628030.png`

Fuente de datos (solo lectura):
- `account.payment` + `l10n_ar_withholding_ids` (account.move.line)
- `account.tax` (para identificar IIBB/SIRCAR y alícuota)

Notas importantes:
- Campos numéricos: ceros a la izquierda.
- Razón social (40): alineada a izquierda, relleno con espacios a derecha.
- El archivo debe tener líneas de **exactamente 163** caracteres + CRLF.
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

DEC3 = Decimal("0.001")
DEC2 = Decimal("0.01")
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


def _yyyymm(iso_date: str) -> str:
    return datetime.strptime(iso_date, "%Y-%m-%d").date().strftime("%Y%m")


def _fmt_centavos(value: Any, width: int) -> str:
    d = Decimal(str(value or 0))
    cents = (abs(d) * 100).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
    n = int(cents) if cents is not None else 0
    return str(n).zfill(width)[-width:]


def _fmt_alicuota_5(value_percent: Any) -> str:
    """
    Campo 12: 5 dígitos, (2 enteros + 3 decimales) sin coma.
    Ej: 2.5% => 02.500 => "02500"
    """
    d = Decimal(str(value_percent or 0)).quantize(DEC3, rounding=ROUND_HALF_UP)
    n = int((d * 1000).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    if n < 0:
        n = -n
    return str(n).zfill(5)[-5:]


def _ascii_upper_pad_right(s: str, width: int) -> str:
    # Normaliza acentos y fuerza ASCII
    s = (s or "").strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if ord(ch) < 128)
    s = s.upper()
    if len(s) >= width:
        return s[:width]
    return s + (" " * (width - len(s)))


def _nro_comp_12(raw: str) -> str:
    # Campo 7: 12 dígitos (nro factura u OP). Si no hay, ceros.
    d = _digits_only(raw or "")
    return d[-12:].zfill(12)


def _pv_num_16(raw: str) -> str:
    """
    Campo 10: 16 dígitos: PV(4) + NUM(12).
    Si viene PV-NUM, lo arma. Si no, rescata dígitos.
    """
    s = (raw or "").strip()
    m = re.match(r"^\\s*(\\d{1,9})\\s*[-/]\\s*(\\d{1,12})\\s*$", s)
    if m:
        pv = int(m.group(1))
        num = int(m.group(2))
        pv = max(0, min(pv, 9999))
        num = max(0, min(num, 999999999999))
        return f"{pv:04d}{num:012d}"
    d = _digits_only(s)
    if not d:
        return "0" * 16
    return d[-16:].zfill(16)


def _tax_ids_iibb(models: Any, db: str, uid: int, pwd: str) -> set[int]:
    dom: list[Any] = [
        "|",
        "|",
        ("l10n_ar_tax_type", "ilike", "iibb"),
        ("name", "ilike", "IIBB"),
        ("name", "ilike", "SIRCAR"),
    ]
    ids: list[int] = models.execute_kw(db, uid, pwd, "account.tax", "search", [dom])
    return {int(x) for x in ids}


def _jurisd_3_desde_tax(tax: dict, fallback: str) -> str:
    # Heurística: últimos 3 dígitos de l10n_ar_code si existen; si no, fallback
    code = str(tax.get("l10n_ar_code") or "")
    d = _digits_only(code)
    if len(d) >= 3:
        return d[-3:].zfill(3)
    return str(fallback).zfill(3)[:3]


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

    tax_iibb = _tax_ids_iibb(models, db, uid, pwd)
    if not tax_iibb:
        raise SystemExit("No hay impuestos IIBB/SIRCAR detectados en account.tax.")

    pay_domain = [
        ("date", ">=", desde),
        ("date", "<=", hasta),
        ("partner_type", "=", "supplier"),
        ("l10n_ar_withholding_ids", "!=", False),
        ("state", "in", ["paid", "in_process", "posted"]),
    ]
    pay_ids: list[int] = models.execute_kw(db, uid, pwd, "account.payment", "search", [pay_domain], {"order": "date asc, id asc"})
    if not pay_ids:
        raise SystemExit("No se encontraron pagos con retenciones en el rango.")

    pays: list[dict] = models.execute_kw(
        db,
        uid,
        pwd,
        "account.payment",
        "read",
        [pay_ids],
        {"fields": ["id", "name", "date", "partner_id", "l10n_ar_withholding_ids", "reconciled_bill_ids"]},
    )

    partner_ids = sorted({p["partner_id"][0] for p in pays if p.get("partner_id")})
    partners: dict[int, dict] = {}
    if partner_ids:
        pr = models.execute_kw(db, uid, pwd, "res.partner", "read", [partner_ids], {"fields": ["id", "name", "vat", "l10n_ar_vat"]})
        partners = {int(r["id"]): r for r in pr}

    # Facturas conciliadas (para extraer nro fiscal y letra si se necesita)
    bill_ids: set[int] = set()
    for p in pays:
        for bt in (p.get("reconciled_bill_ids") or []):
            if isinstance(bt, (list, tuple)) and bt:
                bill_ids.add(int(bt[0]))
            elif isinstance(bt, int):
                bill_ids.add(int(bt))
    bills_map: dict[int, dict] = {}
    if bill_ids:
        br = models.execute_kw(
            db,
            uid,
            pwd,
            "account.move",
            "read",
            [sorted(bill_ids)],
            {"fields": ["id", "name", "invoice_date", "amount_total", "l10n_latam_document_number"]},
        )
        bills_map = {int(r["id"]): r for r in br}

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

    tax_ids = sorted({l["tax_line_id"][0] for l in lines if l.get("tax_line_id")})
    taxes: dict[int, dict] = {}
    if tax_ids:
        tr = models.execute_kw(db, uid, pwd, "account.tax", "read", [tax_ids], {"fields": ["id", "name", "amount", "l10n_ar_tax_type", "l10n_ar_code"]})
        taxes = {int(r["id"]): r for r in tr}

    pays_by_id = {int(p["id"]): p for p in pays}

    out_lines: list[str] = []
    for l in lines:
        pid = l.get("payment_id")
        tid = l.get("tax_line_id")
        if not pid or not tid:
            continue
        pid_i = int(pid[0])
        tid_i = int(tid[0])
        if tid_i not in tax_iibb:
            continue
        tax = taxes.get(tid_i) or {}

        p = pays_by_id.get(pid_i) or {}
        fecha_ret = _ddmmyyyy(str(p["date"]))
        mes_anio = _yyyymm(str(p["date"]))
        cuota = cuota_1 if cuota_1 else ("1" if int(str(p["date"])[8:10]) <= 15 else "2")

        # Nº comprobante: priorizar factura conciliada; si no, pago
        bills = []
        for bt in (p.get("reconciled_bill_ids") or []):
            bid = int(bt[0]) if isinstance(bt, (list, tuple)) and bt else int(bt) if isinstance(bt, int) else None
            if bid and bid in bills_map:
                bills.append(bills_map[bid])
        bills.sort(key=lambda r: (r.get("invoice_date") or "", int(r["id"])))
        bill = bills[0] if bills else None

        raw_12 = str((bill or {}).get("l10n_latam_document_number") or (bill or {}).get("name") or p.get("name") or "")
        campo7_12 = _nro_comp_12(raw_12)
        campo10_16 = _pv_num_16(raw_12)

        # Tipo comprobante (F/L/P) + Letra (A/B/C/esp)
        tipo_comp = "F" if bill else "P"
        letra = " "
        name_hint = str((bill or {}).get("name") or "")
        for cand in ("FA-A", "FC-A", "NC-A", "ND-A"):
            if cand in name_hint:
                letra = "A"
                break
        if letra == " ":
            for cand in ("FA-B", "FC-B", "NC-B", "ND-B"):
                if cand in name_hint:
                    letra = "B"
                    break
        if letra == " ":
            for cand in ("FA-C", "FC-C", "NC-C", "ND-C"):
                if cand in name_hint:
                    letra = "C"
                    break

        base_15 = _fmt_centavos(l.get("tax_base_amount"), 15)
        monto_ret_15 = _fmt_centavos(l.get("credit") or l.get("balance") or 0, 15)
        alic_5 = _fmt_alicuota_5(tax.get("amount"))

        partner_id = int(p["partner_id"][0]) if p.get("partner_id") else 0
        partner = partners.get(partner_id, {})
        cuit_suj = _digits_only(str(partner.get("l10n_ar_vat") or partner.get("vat") or "")).rjust(11, "0")[-11:]
        razon_40 = _ascii_upper_pad_right(str(partner.get("name") or ""), 40)

        # Jurisdicción: obligatoria (agente y sujeto).
        # En Nakel suele ser un valor fijo (ej. 907/920) definido por el estudio.
        juris = str(jurisdiccion_3).zfill(3)[:3]
        juris_suj = str(jurisdiccion_sujeto_3 or juris).zfill(3)[:3]

        # Campo 1: tipo de registro (1=retención)
        tipo_reg = "1"
        tipo_doc_1 = "1"  # CUIT
        otros_11 = "0" * 11

        out = (
            tipo_reg  # 1
            + juris  # 3
            + _digits_only(cuit_agente_11).rjust(11, "0")[-11:]  # 11
            + mes_anio  # 6
            + str(cuota)[:1]  # 1
            + fecha_ret  # 10
            + campo7_12  # 12
            + tipo_comp  # 1
            + letra  # 1
            + campo10_16  # 16
            + base_15  # 15
            + alic_5  # 5
            + monto_ret_15  # 15
            + tipo_doc_1  # 1
            + cuit_suj  # 11
            + juris_suj  # 3
            + razon_40  # 40
            + otros_11  # 11
        )
        if len(out) != LREG:
            raise SystemExit(f"Línea {len(out)} != {LREG} (pid={pid_i}, aml={l['id']})")
        out_lines.append(out)

    if not out_lines:
        raise SystemExit("No se encontraron retenciones IIBB/SIRCAR para exportar (layout 163).")

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
    ap.add_argument("--cuit-agente", required=True, help="CUIT agente (11 dígitos, sin guiones).")
    ap.add_argument(
        "--jurisdiccion",
        default="907",
        help="Jurisdicción agente (3 dígitos). Default 907 (Nakel).",
    )
    ap.add_argument(
        "--jurisdiccion-sujeto",
        default=None,
        help="Jurisdicción del sujeto (3 dígitos). Si se omite, usa la misma que --jurisdiccion.",
    )
    ap.add_argument("--cuota", default=None, help="1 o 2. Si no se informa, se calcula por día (<=15 => 1).")
    ap.add_argument(
        "--out",
        type=Path,
        default=Path(__file__).resolve().parent / "out" / "SIRCAR_163.TXT",
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

