#!/usr/bin/env python3
"""
Planilla XLSX: IIBB sufrida en compras (facturas de proveedor).

Interpretación (Odoo):
- Líneas `account.move.line` publicadas (`parent_state=posted`) con `tax_line_id` IIBB/SIRCAR.
- Se filtra a compras dejando movimientos `account.move.move_type in ('in_invoice', 'in_refund')`.

Salida: exportador-excel/out/retenciones_sufridas_iibb_<desde>_a_<hasta>.xlsx
"""

from __future__ import annotations

import argparse
import csv
import sys
from datetime import datetime
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


def main(argv: Iterable[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--desde", type=_iso, required=True)
    ap.add_argument("--hasta", type=_iso, required=True)
    ap.add_argument("--out-dir", type=Path, default=Path(__file__).resolve().parent / "out")
    args = ap.parse_args(list(argv) if argv is not None else None)

    root = resolve_project_root(Path(__file__).resolve())
    sys.path.insert(0, str(root))
    from nakel_import_paths import prepend_config_nakel_sys_path  # type: ignore

    prepend_config_nakel_sys_path(root)
    from config_nakel import ODOO_CONFIG_MASTER_DEV  # type: ignore

    models, uid = odoo_connect(ODOO_CONFIG_MASTER_DEV)
    db, pwd = ODOO_CONFIG_MASTER_DEV["db"], ODOO_CONFIG_MASTER_DEV["password"]

    tax_dom = ["|", "|", ("l10n_ar_tax_type", "ilike", "iibb"), ("name", "ilike", "IIBB"), ("name", "ilike", "SIRCAR")]
    iibb_tax_ids: list[int] = models.execute_kw(db, uid, pwd, "account.tax", "search", [tax_dom])
    iibb_tax_ids = [int(x) for x in iibb_tax_ids]
    if not iibb_tax_ids:
        raise SystemExit("No se encontraron impuestos IIBB/SIRCAR (por nombre o l10n_ar_tax_type).")

    line_dom = [
        ("date", ">=", args.desde),
        ("date", "<=", args.hasta),
        ("parent_state", "=", "posted"),
        ("tax_line_id", "in", iibb_tax_ids),
    ]
    line_ids: list[int] = models.execute_kw(db, uid, pwd, "account.move.line", "search", [line_dom], {"order": "date asc, id asc"})
    if not line_ids:
        raise SystemExit("No hay líneas IIBB/SIRCAR en el rango dado.")

    lines: list[dict] = models.execute_kw(
        db,
        uid,
        pwd,
        "account.move.line",
        "read",
        [line_ids],
        {"fields": ["id", "date", "move_id", "partner_id", "tax_base_amount", "tax_line_id"]},
    )

    move_ids = sorted({l["move_id"][0] for l in lines if l.get("move_id")})
    moves: dict[int, dict] = {}
    if move_ids:
        mr = models.execute_kw(db, uid, pwd, "account.move", "read", [move_ids], {"fields": ["id", "name", "ref", "move_type"]})
        moves = {int(r["id"]): r for r in mr}

    partner_ids = sorted({l["partner_id"][0] for l in lines if l.get("partner_id")})
    partners: dict[int, dict] = {}
    if partner_ids:
        pr = models.execute_kw(db, uid, pwd, "res.partner", "read", [partner_ids], {"fields": ["id", "name", "vat", "l10n_ar_vat"]})
        partners = {int(r["id"]): r for r in pr}

    tax_ids = sorted({l["tax_line_id"][0] for l in lines if l.get("tax_line_id")})
    taxes: dict[int, dict] = {}
    if tax_ids:
        tr = models.execute_kw(db, uid, pwd, "account.tax", "read", [tax_ids], {"fields": ["id", "name", "amount", "l10n_ar_tax_type"]})
        taxes = {int(r["id"]): r for r in tr}

    out_rows: list[RetencionRow] = []
    for l in sorted(lines, key=lambda r: int(r["id"])):
        mid = l["move_id"][0] if l.get("move_id") else None
        if not mid:
            continue
        move = moves.get(int(mid)) or {}
        if (move.get("move_type") or "") not in ("in_invoice", "in_refund"):
            continue
        tax = taxes.get(l["tax_line_id"][0]) if l.get("tax_line_id") else None
        if not tax or not is_iibb_tax(tax):
            continue

        partner_id = int(l["partner_id"][0]) if l.get("partner_id") else 0
        partner = partners.get(partner_id, {})
        base = abs_money(l.get("tax_base_amount"))
        rate = abs_money(tax.get("amount"))
        importe = calc_importe_from_base_rate(base, tax.get("amount"))

        out_rows.append(
            RetencionRow(
                payment_id=int(mid),
                payment_name=str(move.get("name") or ""),
                payment_date=str(l.get("date") or ""),
                partner_name=str(partner.get("name") or ""),
                partner_cuit=cuit11_from_partner(partner),
                tax_name=str(tax.get("name") or ""),
                tax_rate=rate,
                base=base,
                importe=importe,
                regimen="",
                extra="",
            )
        )

    if not out_rows:
        raise SystemExit(
            "No se encontraron líneas IIBB en compras (in_invoice/in_refund) en el rango. "
            "Si el estudio esperaba otra cosa (ej. retenciones en cobros), hay que confirmar el origen en Odoo."
        )

    out_dir: Path = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    tmp_csv = out_dir / f"retenciones_sufridas_iibb_{args.desde}_a_{args.hasta}.csv"
    out_xlsx = out_dir / f"retenciones_sufridas_iibb_{args.desde}_a_{args.hasta}.xlsx"

    headers = [
        "move_id",
        "move_name",
        "fecha",
        "proveedor_name",
        "proveedor_cuit",
        "tax_name",
        "alicuota_percent",
        "base",
        "importe",
    ]
    with tmp_csv.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(headers)
        for r in out_rows:
            w.writerow(
                [
                    r.payment_id,
                    r.payment_name,
                    ddmmyyyy(r.payment_date),
                    r.partner_name,
                    r.partner_cuit,
                    r.tax_name,
                    f"{r.tax_rate:.2f}",
                    f"{r.base:.2f}",
                    f"{r.importe:.2f}",
                ]
            )

    out = convert_csv_to_xlsx(tmp_csv, out_xlsx)
    print(f"OK: {out}")
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())

