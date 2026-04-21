#!/usr/bin/env python3
"""
Genera un archivo tipo `RET-DGR.TXT` (CSV) consultando Odoo master_dev por XML-RPC.

Objetivo:
- Solo lectura (no escribe nada en Odoo).
- Exporta retenciones (withholdings) asociadas a pagos a proveedores.

Notas:
- En Odoo 18 Argentina, las retenciones están en `account.payment.l10n_ar_withholding_ids`
  como líneas contables (`account.move.line`) con `tax_base_amount`, `balance/credit/debit`
  y el impuesto en `tax_line_id` (`account.tax`).
- Este export filtra por defecto impuestos con `l10n_ar_tax_type` que empiece por "iibb"
  o nombre que contenga "SIRCAR"/"IIBB".
"""

from __future__ import annotations

import argparse
import csv
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


def _d2(x: Any) -> Decimal:
    if x is None or x is False:
        return Decimal("0")
    try:
        return Decimal(str(x))
    except Exception:
        return Decimal("0")


def _abs_money(x: Any) -> Decimal:
    return abs(_d2(x)).quantize(DEC2, rounding=ROUND_HALF_UP)


def _fmt_money_dot(x: Decimal) -> str:
    # CSV del ejemplo usa punto decimal.
    return f"{x:.2f}"


def _fmt_rate(x: Any) -> str:
    # En el ejemplo la alícuota viene con hasta 2 decimales.
    d = _d2(x).quantize(DEC2, rounding=ROUND_HALF_UP)
    s = f"{d:.2f}"
    # Evitar "-0.00"
    return "0.00" if s in ("-0.00", "-0,00") else s


def _calc_importe_from_base_rate(base: Decimal, rate_percent: Any) -> Decimal:
    # Importante: SIAP valida con alícuota a 2 decimales (como se exporta).
    rate = _d2(rate_percent).quantize(DEC2, rounding=ROUND_HALF_UP)
    return (base * rate / Decimal("100")).quantize(DEC2, rounding=ROUND_HALF_UP)


def _digits_only(s: str) -> str:
    return re.sub(r"\D+", "", s or "")

def _normalize_payment_name_no_separators(name: str, *, max_len: int = 16) -> str:
    """
    Normaliza `account.payment.name` para cumplir "sin separadores" (SICORE/ARCA).
    Ej: PGAL1/26-27/0370 -> PGAL126270370
    """
    if not name:
        return ""
    out = re.sub(r"[^A-Za-z0-9]+", "", str(name))
    return out[:max_len]


def _cuit_11_from_partner(row: dict) -> str:
    # Prioriza l10n_ar_vat (Argentina) y cae a vat.
    raw = row.get("l10n_ar_vat") or row.get("vat") or ""
    dig = _digits_only(str(raw))
    if len(dig) == 11:
        return dig
    # si viene como 80+CUIT o con prefijos raros, intentar rescatar últimos 11
    if len(dig) > 11:
        return dig[-11:]
    return dig.rjust(11, "0") if dig else ""


def _ddmmyyyy(iso_date: str) -> str:
    # Odoo suele devolver YYYY-MM-DD
    dt = datetime.strptime(iso_date, "%Y-%m-%d").date()
    return dt.strftime("%d/%m/%Y")


def odoo_connect(cfg: dict) -> tuple[Any, int]:
    url = cfg["url"].rstrip("/")
    common = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/common", allow_none=True)
    uid = common.authenticate(cfg["db"], cfg["username"], cfg["password"], {})
    if not uid:
        raise SystemExit("Autenticación Odoo fallida (revisá ODOO_CONFIG_MASTER_DEV).")
    return xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/object", allow_none=True), int(uid)


@dataclass(frozen=True)
class RetRow:
    nro_registro: str
    lote: str
    sublote: str
    orden_pago: str
    cuit: str
    fecha: str
    base: str
    alicuota: str
    importe: str
    codigo_regimen: str
    codigo_extra: str

    def to_csv_row(self) -> list[str]:
        return [
            self.nro_registro,
            self.lote,
            self.sublote,
            self.orden_pago,
            self.cuit,
            self.fecha,
            self.base,
            self.alicuota,
            self.importe,
            self.codigo_regimen,
            self.codigo_extra,
        ]


def _is_dgr_tax(tax: dict, prefer_iibb_only: bool) -> bool:
    nm = (tax.get("name") or "").upper()
    tt = (tax.get("l10n_ar_tax_type") or "").lower()
    if prefer_iibb_only:
        return tt.startswith("iibb") or ("SIRCAR" in nm) or ("IIBB" in nm)
    # modo amplio: cualquier withholding argentino (si quisieras Ganancias también)
    return True


def _alloc_by_weights(total: Decimal, weights: list[Decimal]) -> list[Decimal]:
    """
    Prorratea `total` según `weights` (>=0), manteniendo 2 decimales y garantizando
    que la suma final sea exactamente `total` (ajustando el último ítem).
    """
    if not weights:
        return []
    s = sum(weights)
    if s <= Decimal("0"):
        # si no hay pesos, todo al primero
        out = [Decimal("0.00")] * len(weights)
        out[0] = total.quantize(DEC2, rounding=ROUND_HALF_UP)
        return out
    out: list[Decimal] = []
    acc = Decimal("0.00")
    for i, w in enumerate(weights):
        if i == len(weights) - 1:
            v = (total - acc).quantize(DEC2, rounding=ROUND_HALF_UP)
        else:
            v = (total * w / s).quantize(DEC2, rounding=ROUND_HALF_UP)
            acc += v
        out.append(v)
    return out


def generar(
    desde: str,
    hasta: str,
    *,
    out_path: Path,
    codigo_regimen: str,
    codigo_extra: str,
    prefer_iibb_only: bool,
    usar_payment_name_normalizado: bool,
    por_factura: bool,
) -> Path:
    cfg = ODOO_CONFIG_MASTER_DEV
    models, uid = odoo_connect(cfg)
    db, pwd = cfg["db"], cfg["password"]

    pay_domain = [
        ("date", ">=", desde),
        ("date", "<=", hasta),
        ("partner_type", "=", "supplier"),
        ("l10n_ar_withholding_ids", "!=", False),
    ]
    pay_ids: list[int] = models.execute_kw(
        db,
        uid,
        pwd,
        "account.payment",
        "search",
        [pay_domain],
        {"order": "date asc, id asc"},
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
        {"fields": ["id", "name", "date", "partner_id", "l10n_ar_withholding_ids", "reconciled_bill_ids"]},
    )

    # Partners (para CUIT)
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
            {"fields": ["id", "vat", "l10n_ar_vat", "name"]},
        )
        partners = {int(r["id"]): r for r in pr}

    # Withholding move lines
    line_ids = sorted(
        {lid for p in pays for lid in (p.get("l10n_ar_withholding_ids") or [])}
    )
    lines: list[dict] = models.execute_kw(
        db,
        uid,
        pwd,
        "account.move.line",
        "read",
        [line_ids],
        {
            "fields": [
                "id",
                "date",
                "payment_id",
                "partner_id",
                "tax_base_amount",
                "balance",
                "debit",
                "credit",
                "tax_line_id",
            ]
        },
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
            {
                "fields": [
                    "id",
                    "name",
                    "amount",
                    "amount_type",
                    "l10n_ar_code",
                    "l10n_ar_tax_type",
                    "l10n_ar_state_id",
                ]
            },
        )
        taxes = {int(r["id"]): r for r in tr}

    # Bills (facturas proveedor) reconciliadas contra los pagos del rango.
    bill_ids = sorted({bid for p in pays for bid in (p.get("reconciled_bill_ids") or [])})
    bills_by_id: dict[int, dict] = {}
    if bill_ids:
        br = models.execute_kw(
            db,
            uid,
            pwd,
            "account.move",
            "read",
            [bill_ids],
            {"fields": ["id", "move_type", "amount_total", "ref", "name", "l10n_latam_document_number"]},
        )
        bills_by_id = {int(b["id"]): b for b in br}

    # Agrupar por payment, filtrando impuestos DGR/SIRCAR (IIBB)
    lines_by_payment: dict[int, list[dict]] = {}
    for l in lines:
        pid = l["payment_id"][0] if l.get("payment_id") else None
        if not pid:
            continue
        tax = taxes.get(l["tax_line_id"][0]) if l.get("tax_line_id") else None
        if not tax:
            continue
        if not _is_dgr_tax(tax, prefer_iibb_only=prefer_iibb_only):
            continue
        lines_by_payment.setdefault(int(pid), []).append(l)

    filas: list[RetRow] = []
    reg = 0
    for p in sorted(pays, key=lambda r: (r.get("date") or "", int(r["id"]))):
        pid = int(p["id"])
        wlines = lines_by_payment.get(pid) or []
        if not wlines:
            continue

        partner_id = p["partner_id"][0] if p.get("partner_id") else None
        partner = partners.get(int(partner_id)) if partner_id else {}
        cuit = _cuit_11_from_partner(partner or {})

        fecha = _ddmmyyyy(p["date"])
        if usar_payment_name_normalizado:
            orden_pago = _normalize_payment_name_no_separators(p.get("name") or "")
        else:
            orden_pago = str(pid).zfill(12)

        bill_list = [bills_by_id.get(int(bid), {}) for bid in (p.get("reconciled_bill_ids") or [])]
        bill_list = [b for b in bill_list if b and (b.get("move_type") in ("in_invoice", "in_refund"))]
        weights = [_abs_money(b.get("amount_total")) for b in bill_list]

        for l in sorted(wlines, key=lambda r: int(r["id"])):
            base = _abs_money(l.get("tax_base_amount"))
            # Monto retenido: tomar crédito/debito/balance como absoluto
            importe = _abs_money(l.get("credit") or l.get("balance") or 0)  # suele venir en credit
            if importe == Decimal("0.00"):
                # fallback por si viene solo en balance
                importe = _abs_money(l.get("balance"))

            tax = taxes.get(l["tax_line_id"][0]) if l.get("tax_line_id") else {}
            alicuota = _fmt_rate(tax.get("amount"))

            if por_factura and bill_list:
                base_parts = _alloc_by_weights(base, weights)
                for bp in base_parts:
                    # Importante: SIAP valida que retenido == base * alicuota / 100.
                    ip = _calc_importe_from_base_rate(bp, tax.get("amount"))
                    reg += 1
                    nro_registro = str(reg).zfill(5)
                    filas.append(
                        RetRow(
                            nro_registro=nro_registro,
                            lote="1",
                            sublote="1",
                            orden_pago=orden_pago,
                            cuit=cuit,
                            fecha=fecha,
                            base=_fmt_money_dot(bp),
                            alicuota=alicuota,
                            importe=_fmt_money_dot(ip),
                            codigo_regimen=codigo_regimen,
                            codigo_extra=codigo_extra,
                        )
                    )
            else:
                reg += 1
                nro_registro = str(reg).zfill(5)
                filas.append(
                    RetRow(
                        nro_registro=nro_registro,
                        lote="1",
                        sublote="1",
                        orden_pago=orden_pago,
                        cuit=cuit,
                        fecha=fecha,
                        base=_fmt_money_dot(base),
                        alicuota=alicuota,
                        importe=_fmt_money_dot(importe),
                        codigo_regimen=codigo_regimen,
                        codigo_extra=codigo_extra,
                    )
                )

    if not filas:
        raise SystemExit(
            "Se encontraron pagos con retenciones, pero ninguna coincidió con el filtro DGR/IIBB."
        )

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f, delimiter=",", lineterminator="\r\n")
        for row in filas:
            w.writerow(row.to_csv_row())

    return out_path


def _iso(s: str) -> str:
    datetime.strptime(s, "%Y-%m-%d")
    return s


def _default_range_mes_actual() -> tuple[str, str]:
    hoy = date.today()
    desde = hoy.replace(day=1)
    if desde.month == 12:
        hasta = desde.replace(year=desde.year + 1, month=1, day=1)  # next month first day
    else:
        hasta = desde.replace(month=desde.month + 1, day=1)
    # hasta inclusive = último día del mes actual
    # (último día = primer día del mes siguiente - 1 día)
    from datetime import timedelta

    hasta_inc = hasta - timedelta(days=1)
    return desde.isoformat(), hasta_inc.isoformat()


def main(argv: Iterable[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    d_desde, d_hasta = _default_range_mes_actual()
    ap.add_argument("--desde", type=_iso, default=d_desde, help="YYYY-MM-DD (incl.)")
    ap.add_argument("--hasta", type=_iso, default=d_hasta, help="YYYY-MM-DD (incl.)")
    ap.add_argument(
        "--out",
        type=Path,
        default=Path(__file__).resolve().parent.parent / "out" / "RET-DGR.TXT",
        help="Ruta del TXT a generar",
    )
    ap.add_argument("--codigo-regimen", default="001", help="Columna 10 (ej. 001)")
    ap.add_argument("--codigo-extra", default="907", help="Columna 11 (ej. 907)")
    ap.add_argument(
        "--incluir-no-iibb",
        action="store_true",
        help="Si se activa, incluye cualquier retención (no solo IIBB/SIRCAR).",
    )
    ap.add_argument(
        "--orden-por-payment-id",
        action="store_true",
        default=False,
        help="Usa payment.id zfill(12) como columna 4 (en vez de payment.name normalizado).",
    )
    ap.add_argument(
        "--por-factura",
        action="store_true",
        default=True,
        help="Desagrega por factura reconciliada (prorratea base/importe). Default: activo.",
    )
    ap.add_argument(
        "--por-pago",
        action="store_true",
        default=False,
        help="Fuerza modo 1 línea por pago (sin desagregar facturas).",
    )
    args = ap.parse_args(list(argv) if argv is not None else None)

    out = generar(
        args.desde,
        args.hasta,
        out_path=args.out,
        codigo_regimen=str(args.codigo_regimen).strip(),
        codigo_extra=str(args.codigo_extra).strip(),
        prefer_iibb_only=not args.incluir_no_iibb,
        usar_payment_name_normalizado=not bool(args.orden_por_payment_id),
        por_factura=bool(args.por_factura) and (not bool(args.por_pago)),
    )
    print(f"OK: {out}")
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())

