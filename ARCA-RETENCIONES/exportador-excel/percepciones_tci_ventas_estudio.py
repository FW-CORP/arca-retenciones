#!/usr/bin/env python3
"""
Planilla tipo estudio: **percepción TCI (Tasa de Comercio e Industria) multilateral**
practicada en **ventas** (facturas de cliente).

Filtra solo líneas cuyo `tax_line_id` coincide con impuestos de venta cuyo nombre
contiene el patrón configurado (por defecto: **Percepcion MCR TCI Multilateral**,
p. ej. el 2% en `master_dev`; no incluye «TCI sin Constancia» salvo que cambies el patrón).

Columnas (misma grilla que `percepciones_iibb_ventas_estudio.py`):
- Nro.Inscripcion, Cuit, Neto Gravado, Fecha, Percepcion, Nro.Percepcion (nº fiscal)

Salida: `exportador-excel/out/percepciones_tci_ventas_estudio_<desde>_a_<hasta>.xlsx`
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Iterable

from _odoo_retenciones import digits_only, odoo_connect, resolve_project_root
from _xlsx import convert_csv_to_xlsx

DEC2 = Decimal("0.01")


def _iso(s: str) -> str:
    datetime.strptime(s, "%Y-%m-%d")
    return s


def _tax_ids_por_nombre_venta(
    models: Any, db: str, uid: int, pwd: str, nombre_ilike: str
) -> set[int]:
    pat = (nombre_ilike or "").strip()
    if not pat:
        raise SystemExit("El patrón de nombre de impuesto no puede estar vacío.")
    dom: list[Any] = [
        "&",
        ("name", "ilike", pat),
        ("type_tax_use", "=", "sale"),
    ]
    ids = models.execute_kw(db, uid, pwd, "account.tax", "search", [dom])
    return {int(x) for x in ids}


def _d2(x: Any) -> Decimal:
    try:
        return Decimal(str(x or 0))
    except Exception:
        return Decimal("0")


def _raw_latam_nro(inv: dict) -> str:
    raw = str(inv.get("l10n_latam_document_number") or "").strip()
    if not raw:
        raw = str(inv.get("name") or "").strip()
        m = re.search(r"\b(\d{4,5}\s*[-/]\s*\d{6,8})\b", raw)
        if m:
            raw = m.group(1)
    return raw


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


def _nro_factura_columna(inv: dict, formato: str) -> str:
    raw = _raw_latam_nro(inv)
    if formato == "pv_guion_num":
        parsed = _fiscal_pv_num_from_raw(raw)
        if parsed:
            pv, num = parsed
            return f"{pv:04d}-{num:08d}"
        d = digits_only(raw)
        if len(d) >= 12:
            try:
                return f"{int(d[:4]):04d}-{int(d[4:12]):08d}"
            except Exception:
                pass
        return d
    return digits_only(raw)


def _mdyyyy(iso_date: str) -> str:
    dt = datetime.strptime(iso_date[:10], "%Y-%m-%d").date()
    return f"{dt.month}/{dt.day}/{dt.year}"


def _base_perc_signed(move_type: str, tax_base_amount: Any, balance: Any) -> tuple[Decimal, Decimal]:
    base = _d2(tax_base_amount).quantize(DEC2, rounding=ROUND_HALF_UP)
    bal = _d2(balance)
    if move_type == "out_refund":
        return (-base.copy_abs(), -bal)
    return (base.copy_abs(), -bal)


def main(argv: Iterable[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--desde", type=_iso, required=True)
    ap.add_argument("--hasta", type=_iso, required=True)
    ap.add_argument("--out-dir", type=Path, default=Path(__file__).resolve().parent / "out")
    ap.add_argument(
        "--impuesto-nombre-contiene",
        default="Percepcion MCR TCI Multilateral",
        help="Subcadena ilike sobre account.tax.name (venta). Por defecto solo multilateral TCI.",
    )
    ap.add_argument(
        "--layout-estudio",
        action="store_true",
        help="Anteponer 3 filas vacías como en la planilla ejemplo del estudio.",
    )
    ap.add_argument(
        "--nro-factura-formato",
        choices=("solo_digitos", "pv_guion_num"),
        default="solo_digitos",
        help="Columna Nro.Percepcion: solo dígitos o PV-número.",
    )
    args = ap.parse_args(list(argv) if argv is not None else None)

    root = resolve_project_root(Path(__file__).resolve())
    sys.path.insert(0, str(root))
    from nakel_import_paths import prepend_config_nakel_sys_path  # type: ignore

    prepend_config_nakel_sys_path(root)
    from config_nakel import ODOO_CONFIG_MASTER_DEV  # type: ignore

    models, uid = odoo_connect(ODOO_CONFIG_MASTER_DEV)
    db, pwd = ODOO_CONFIG_MASTER_DEV["db"], ODOO_CONFIG_MASTER_DEV["password"]

    tax_ids = _tax_ids_por_nombre_venta(models, db, uid, pwd, args.impuesto_nombre_contiene)
    if not tax_ids:
        raise SystemExit(
            f"No hay impuestos de venta (`type_tax_use=sale`) cuyo nombre contenga "
            f"{args.impuesto_nombre_contiene!r}. Revisar maestro de impuestos o el patrón."
        )

    move_dom = [
        ("move_type", "in", ["out_invoice", "out_refund", "out_debit"]),
        ("state", "=", "posted"),
        ("invoice_date", ">=", args.desde),
        ("invoice_date", "<=", args.hasta),
    ]
    move_ids: list[int] = models.execute_kw(db, uid, pwd, "account.move", "search", [move_dom], {"order": "invoice_date asc, id asc"})
    if not move_ids:
        raise SystemExit("No hay facturas de cliente publicadas en el rango.")

    line_dom = [
        ("move_id", "in", move_ids),
        ("tax_line_id", "in", sorted(tax_ids)),
    ]
    line_ids: list[int] = models.execute_kw(
        db, uid, pwd, "account.move.line", "search", [line_dom], {"order": "move_id asc, id asc"}
    )
    if not line_ids:
        raise SystemExit(
            "Hay facturas en el rango pero ninguna línea de impuesto TCI con el patrón indicado."
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
            {"fields": ["id", "vat", "l10n_ar_vat", "l10n_ar_gross_income_number"]},
        )
        partners = {int(r["id"]): r for r in pr}

    out_dir: Path = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = f"percepciones_tci_ventas_estudio_{args.desde}_a_{args.hasta}"
    tmp_csv = out_dir / f"{stem}.csv"
    out_xlsx = out_dir / f"{stem}.xlsx"

    headers = ["Nro.Inscripcion", "Cuit", "Neto Gravado", "Fecha", "Percepcion", "Nro.Percepcion"]
    with tmp_csv.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        if args.layout_estudio:
            w.writerow([])
            w.writerow([])
            w.writerow([])
        w.writerow(headers)
        for l in lines:
            mid = l.get("move_id")
            if not mid:
                continue
            inv = inv_map.get(int(mid[0]))
            if not inv:
                continue
            move_type = str(inv.get("move_type") or "")
            inv_date = str(inv.get("invoice_date") or "")[:10]
            if not inv_date:
                continue

            base_v, perc_v = _base_perc_signed(move_type, l.get("tax_base_amount"), l.get("balance"))
            if perc_v == 0 and (l.get("credit") or l.get("debit")):
                bal2 = _d2(l.get("credit")) - _d2(l.get("debit"))
                _, perc_v = _base_perc_signed(move_type, l.get("tax_base_amount"), bal2)

            pid = int(inv["partner_id"][0]) if inv.get("partner_id") else 0
            partner = partners.get(pid, {})
            raw_cuit = str(partner.get("l10n_ar_vat") or partner.get("vat") or "")
            cuit = digits_only(raw_cuit)
            if len(cuit) > 11:
                cuit = cuit[-11:]
            inscr = str(partner.get("l10n_ar_gross_income_number") or "").strip()
            nro_fact = _nro_factura_columna(inv, args.nro_factura_formato)

            w.writerow(
                [
                    inscr,
                    cuit,
                    f"{base_v:.2f}",
                    _mdyyyy(inv_date),
                    f"{perc_v:.2f}",
                    nro_fact,
                ]
            )

    out = convert_csv_to_xlsx(tmp_csv, out_xlsx)
    print(f"OK: {out}")
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
