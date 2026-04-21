#!/usr/bin/env python3
"""
Planilla XLSX: Retenciones de Ganancias (RGAN_CPA) por operación (factura).

Fuente: Odoo master_dev (solo lectura) por XML-RPC.
Criterio:
- Pagos a proveedores con `l10n_ar_withholding_ids`.
- Se filtra a impuestos Ganancias: `account.tax.l10n_ar_tax_type = earnings`.
- Se desagrega por factura reconciliada (`account.payment.reconciled_bill_ids`) y se prorratea
  base/retención según `amount_total` de cada factura (igual criterio que `RGAN_CPA_v2.TXT`).

Salida: exportador-excel/out/retenciones_ganancias_rgan_cpa_<desde>_a_<hasta>.xlsx
"""

from __future__ import annotations

import argparse
import csv
import sys
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Iterable

from _odoo_retenciones import cuit11_from_partner, odoo_connect, resolve_project_root
from _xlsx import convert_csv_to_xlsx


DEC2 = Decimal("0.01")


def _iso(s: str) -> str:
    datetime.strptime(s, "%Y-%m-%d")
    return s


def _d2(x: Any) -> Decimal:
    if x is None or x is False:
        return Decimal("0")
    try:
        return Decimal(str(x))
    except Exception:
        return Decimal("0")


def _abs_money(x: Any) -> Decimal:
    return abs(_d2(x)).quantize(DEC2, rounding=ROUND_HALF_UP)


def _ddmmyyyy(iso_date: str) -> str:
    dt = datetime.strptime(iso_date, "%Y-%m-%d").date()
    return dt.strftime("%d/%m/%Y")


def _alloc_by_weights(total: Decimal, weights: list[Decimal]) -> list[Decimal]:
    if not weights:
        return []
    s = sum(weights)
    if s <= Decimal("0"):
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


def _calc_alicuota_percent(base: Decimal, reten: Decimal) -> Decimal:
    if base == Decimal("0.00"):
        return Decimal("0.00")
    return (reten * Decimal("100") / base).quantize(DEC2, rounding=ROUND_HALF_UP)


def main(argv: Iterable[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--desde", type=_iso, required=True)
    ap.add_argument("--hasta", type=_iso, required=True)
    ap.add_argument("--rg", default="010", help="Columna RG (default: 010, como en RGAN_CPA)")
    ap.add_argument("--out-dir", type=Path, default=Path(__file__).resolve().parent / "out")
    args = ap.parse_args(list(argv) if argv is not None else None)

    root = resolve_project_root(Path(__file__).resolve())
    sys.path.insert(0, str(root))
    from nakel_import_paths import prepend_config_nakel_sys_path  # type: ignore

    prepend_config_nakel_sys_path(root)
    from config_nakel import ODOO_CONFIG_MASTER_DEV  # type: ignore

    models, uid = odoo_connect(ODOO_CONFIG_MASTER_DEV)
    db, pwd = ODOO_CONFIG_MASTER_DEV["db"], ODOO_CONFIG_MASTER_DEV["password"]

    earn_tax_ids: list[int] = models.execute_kw(
        db,
        uid,
        pwd,
        "account.tax",
        "search",
        [[("l10n_ar_tax_type", "=", "earnings")]],
    )
    earn_tax_ids = [int(x) for x in earn_tax_ids]
    if not earn_tax_ids:
        raise SystemExit("No hay impuestos con l10n_ar_tax_type='earnings' (Ganancias).")
    earn_set = set(earn_tax_ids)

    pay_domain = [
        ("date", ">=", args.desde),
        ("date", "<=", args.hasta),
        ("partner_type", "=", "supplier"),
        ("state", "!=", "cancelled"),
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
    pays_by_id = {int(p["id"]): p for p in pays}

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
            {"fields": ["id", "name", "vat", "l10n_ar_vat"]},
        )
        partners = {int(r["id"]): r for r in pr}

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

    # filtrar a earnings y agrupar por payment
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

    out_dir: Path = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    tmp_csv = out_dir / f"retenciones_ganancias_rgan_cpa_{args.desde}_a_{args.hasta}.csv"
    out_xlsx = out_dir / f"retenciones_ganancias_rgan_cpa_{args.desde}_a_{args.hasta}.xlsx"

    headers = [
        "fecha",
        "proveedor",
        "cuit",
        "base_imponible_ganancias",
        "rg",
        "alicuota_percent",
        "importe_retencion",
    ]

    tot_base = Decimal("0.00")
    tot_ret = Decimal("0.00")

    with tmp_csv.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(headers)

        for p in sorted(pays, key=lambda r: (r.get("date") or "", int(r["id"]))):
            pid = int(p["id"])
            wlines = lines_by_payment.get(pid) or []
            if not wlines:
                continue

            partner_id = int(p["partner_id"][0]) if p.get("partner_id") else 0
            partner = partners.get(partner_id, {})
            prov_name = str(partner.get("name") or "")
            prov_cuit = cuit11_from_partner(partner)
            fecha = _ddmmyyyy(str(p.get("date") or ""))

            bill_list = [bills_by_id.get(int(bid), {}) for bid in (p.get("reconciled_bill_ids") or [])]
            bill_list = [b for b in bill_list if b and (b.get("move_type") in ("in_invoice", "in_refund"))]
            weights = [_abs_money(b.get("amount_total")) for b in bill_list]

            for l in sorted(wlines, key=lambda r: int(r["id"])):
                base_total = _abs_money(l.get("tax_base_amount"))
                ret_total = _abs_money(l.get("credit") or l.get("balance") or 0)
                if ret_total == Decimal("0.00"):
                    ret_total = _abs_money(l.get("balance"))

                if bill_list:
                    base_parts = _alloc_by_weights(base_total, weights)
                    ret_parts = _alloc_by_weights(ret_total, weights)
                    for bp, rp in zip(base_parts, ret_parts, strict=False):
                        tot_base += bp
                        tot_ret += rp
                        w.writerow(
                            [
                                fecha,
                                prov_name,
                                prov_cuit,
                                f"{bp:.2f}",
                                str(args.rg).zfill(3)[:3],
                                f"{_calc_alicuota_percent(bp, rp):.2f}",
                                f"{rp:.2f}",
                            ]
                        )
                else:
                    tot_base += base_total
                    tot_ret += ret_total
                    w.writerow(
                        [
                            fecha,
                            prov_name,
                            prov_cuit,
                            f"{base_total:.2f}",
                            str(args.rg).zfill(3)[:3],
                            f"{_calc_alicuota_percent(base_total, ret_total):.2f}",
                            f"{ret_total:.2f}",
                        ]
                    )

        # Totales
        w.writerow(["", "TOTALES", "", f"{tot_base:.2f}", "", "", f"{tot_ret:.2f}"])

    out = convert_csv_to_xlsx(tmp_csv, out_xlsx)
    print(f"OK: {out}")
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())

