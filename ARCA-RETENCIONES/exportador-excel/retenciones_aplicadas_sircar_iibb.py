#!/usr/bin/env python3
"""
Planilla XLSX: retenciones aplicadas (SIRCAR / IIBB) en pagos a proveedores.

Fuente: Odoo master_dev (solo lectura), mismo criterio que `RET-DGR-SIRCAR.TXT`.
Salida: exportador-excel/out/retenciones_aplicadas_sircar_iibb_<desde>_a_<hasta>.xlsx
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Iterable

from _odoo_retenciones import (
    RetencionRow,
    abs_money,
    calc_importe_from_base_rate,
    cuit11_from_partner,
    ddmmyyyy,
    is_iibb_tax,
    odoo_connect,
    resolve_project_root,
)
from _xlsx import convert_csv_to_xlsx


def _iso(s: str) -> str:
    datetime.strptime(s, "%Y-%m-%d")
    return s


_PV_NRO_RE = re.compile(r"^\s*(\d{1,5})\s*-\s*(\d{1,10})\s*$")


def _fmt_pv_nro(s: str) -> str:
    """
    Formatea un comprobante como PV-NRO, ej: 11-131946 -> 0011-00131946.
    Si no matchea, devuelve string vacío.
    """
    m = _PV_NRO_RE.match((s or "").strip())
    if not m:
        return ""
    pv = m.group(1).zfill(4)
    nro = m.group(2).zfill(8)
    return f"{pv}-{nro}"


def _comprobantes_from_bills(bills: list[dict]) -> str:
    vals: list[str] = []
    for b in bills:
        ref = str(b.get("ref") or "").strip()
        doc = str(b.get("l10n_latam_document_number") or "").strip()
        name = str(b.get("name") or "").strip()
        for cand in (ref, doc, name):
            fmt = _fmt_pv_nro(cand)
            if fmt:
                vals.append(fmt)
                break
    # dedupe manteniendo orden
    out: list[str] = []
    seen: set[str] = set()
    for v in vals:
        if v not in seen:
            seen.add(v)
            out.append(v)
    return ";".join(out)


def _alloc_by_weights(total: Decimal, weights: list[Decimal]) -> list[Decimal]:
    if not weights:
        return []
    s = sum(weights)
    if s <= Decimal("0"):
        out = [Decimal("0.00")] * len(weights)
        out[0] = total
        return out
    out: list[Decimal] = []
    acc = Decimal("0.00")
    for i, w in enumerate(weights):
        if i == len(weights) - 1:
            v = total - acc
        else:
            v = (total * w / s).quantize(Decimal("0.01"))
            acc += v
        out.append(v)
    return out


def main(argv: Iterable[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--desde", type=_iso, required=True)
    ap.add_argument("--hasta", type=_iso, required=True)
    ap.add_argument("--codigo-regimen", default="001")
    ap.add_argument("--codigo-extra", default="907")
    ap.add_argument("--out-dir", type=Path, default=Path(__file__).resolve().parent / "out")
    args = ap.parse_args(list(argv) if argv is not None else None)

    root = resolve_project_root(Path(__file__).resolve())
    sys.path.insert(0, str(root))
    from nakel_import_paths import prepend_config_nakel_sys_path  # type: ignore

    prepend_config_nakel_sys_path(root)
    from config_nakel import ODOO_CONFIG_MASTER_DEV  # type: ignore

    models, uid = odoo_connect(ODOO_CONFIG_MASTER_DEV)
    db, pwd = ODOO_CONFIG_MASTER_DEV["db"], ODOO_CONFIG_MASTER_DEV["password"]

    pay_domain = [
        ("date", ">=", args.desde),
        ("date", "<=", args.hasta),
        ("partner_type", "=", "supplier"),
        ("state", "!=", "cancelled"),
        ("l10n_ar_withholding_ids", "!=", False),
    ]
    pay_ids: list[int] = models.execute_kw(db, uid, pwd, "account.payment", "search", [pay_domain], {"order": "date asc, id asc"})
    if not pay_ids:
        raise SystemExit("No se encontraron pagos a proveedores con retenciones IIBB/SIRCAR en el rango.")

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

    line_ids = sorted({lid for p in pays for lid in (p.get("l10n_ar_withholding_ids") or [])})
    lines: list[dict] = models.execute_kw(
        db,
        uid,
        pwd,
        "account.move.line",
        "read",
        [line_ids],
        {"fields": ["id", "payment_id", "tax_base_amount", "tax_line_id"]},
    )

    tax_ids = sorted({l["tax_line_id"][0] for l in lines if l.get("tax_line_id")})
    taxes: dict[int, dict] = {}
    if tax_ids:
        tr = models.execute_kw(db, uid, pwd, "account.tax", "read", [tax_ids], {"fields": ["id", "name", "amount", "l10n_ar_tax_type"]})
        taxes = {int(r["id"]): r for r in tr}

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
            {"fields": ["id", "move_type", "name", "ref", "l10n_latam_document_number"]},
        )
        bills_by_id = {int(b["id"]): b for b in br}

    pays_by_id = {int(p["id"]): p for p in pays}

    # bills reconciliadas, para desagregar por operación (factura)
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

    out_rows: list[tuple[RetencionRow, str]] = []  # (row, comprobante_pv_nro)
    for l in sorted(lines, key=lambda r: int(r["id"])):
        pid = l["payment_id"][0] if l.get("payment_id") else None
        if not pid:
            continue
        pay = pays_by_id.get(int(pid))
        if not pay:
            continue
        tax = taxes.get(l["tax_line_id"][0]) if l.get("tax_line_id") else None
        if not tax or not is_iibb_tax(tax):
            continue

        partner_id = int(pay["partner_id"][0]) if pay.get("partner_id") else 0
        partner = partners.get(partner_id, {})
        base_total = abs_money(l.get("tax_base_amount"))
        # alícuota a 2 decimales (como se exporta)
        rate = abs_money(tax.get("amount"))

        bill_list = [bills_by_id.get(int(bid), {}) for bid in (pay.get("reconciled_bill_ids") or [])]
        bill_list = [b for b in bill_list if b and (b.get("move_type") in ("in_invoice", "in_refund"))]
        if bill_list:
            weights = [abs_money(b.get("amount_total")) for b in bill_list]
            base_parts = _alloc_by_weights(base_total, weights)
            for b, bp in zip(bill_list, base_parts, strict=False):
                # para consistencia con validaciones: importe = base * alícuota / 100 (con alícuota 2 dec)
                ip = (bp * rate / Decimal("100")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                out_rows.append(
                    (
                        RetencionRow(
                            payment_id=int(pay["id"]),
                            payment_name=str(pay.get("name") or ""),
                            payment_date=str(pay.get("date") or ""),
                            partner_name=str(partner.get("name") or ""),
                            partner_cuit=cuit11_from_partner(partner),
                            tax_name=str(tax.get("name") or ""),
                            tax_rate=rate,
                            base=bp,
                            importe=ip,
                            regimen=str(args.codigo_regimen).strip(),
                            extra=str(args.codigo_extra).strip(),
                        ),
                        _comprobantes_from_bills([b]),
                    )
                )
        else:
            importe_total = (base_total * rate / Decimal("100")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            out_rows.append(
                (
                    RetencionRow(
                        payment_id=int(pay["id"]),
                        payment_name=str(pay.get("name") or ""),
                        payment_date=str(pay.get("date") or ""),
                        partner_name=str(partner.get("name") or ""),
                        partner_cuit=cuit11_from_partner(partner),
                        tax_name=str(tax.get("name") or ""),
                        tax_rate=rate,
                        base=base_total,
                        importe=importe_total,
                        regimen=str(args.codigo_regimen).strip(),
                        extra=str(args.codigo_extra).strip(),
                    ),
                    "",
                )
            )

    if not out_rows:
        raise SystemExit("No se encontraron líneas de retención IIBB/SIRCAR en el rango (tras filtrar taxes).")

    out_dir: Path = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    tmp_csv = out_dir / f"retenciones_aplicadas_sircar_iibb_{args.desde}_a_{args.hasta}.csv"
    out_xlsx = out_dir / f"retenciones_aplicadas_sircar_iibb_{args.desde}_a_{args.hasta}.xlsx"

    headers = [
        "fecha",
        "proveedor",
        "cuit",
        "comprobante_pv_nro",
        "base_imponible",
        "alicuota_percent",
        "monto_retenido",
    ]
    with tmp_csv.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(headers)
        tot_base = Decimal("0.00")
        tot_imp = Decimal("0.00")
        for r, comp in out_rows:
            tot_base += r.base
            tot_imp += r.importe
            w.writerow(
                [
                    ddmmyyyy(r.payment_date),
                    r.partner_name,
                    r.partner_cuit,
                    comp,
                    f"{r.base:.2f}",
                    f"{r.tax_rate:.2f}",
                    f"{r.importe:.2f}",
                ]
            )
        # Fila de totales (para copiar/pegar y validar rápido)
        w.writerow(["", "TOTALES", "", "", f"{tot_base:.2f}", "", f"{tot_imp:.2f}"])

    out = convert_csv_to_xlsx(tmp_csv, out_xlsx)
    print(f"OK: {out}")
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())

