#!/usr/bin/env python3
"""
Genera `RET_IIBB_odoo.xlsx` — libro mayor (formato tipo `RETIIBB.xlsx`)
con movimientos de **retención IIBB / SIRCAR** desde Odoo `master_dev` (XML-RPC).

Criterio:
- Líneas contables `account.move.line` con `tax_line_id` perteneciente a impuestos IIBB/SIRCAR.
  Se consideran IIBB si:
  - `account.tax.l10n_ar_tax_type` contiene "iibb" (ej. iibb_*), o
  - `account.tax.name` contiene "IIBB" o "SIRCAR".
- Solo asientos publicados (`parent_state == 'posted'`).

Estructura (plantilla):
  Fila 2: "Listado Mayor de Cuentas"
  Fila 4: encabezados: Fecha, Transaccion, Detalle, Observaciones, Debe, Haber
  Fila 5: "S.Ini." con debe/haber acumulados antes de `--desde`
  Filas siguientes: detalle

Columnas:
  A Fecha          → fecha de la línea (dd/mm/aaaa)
  B Transacción    → id de `account.move.line`
  C Detalle        → etiqueta de línea +, si hay pago, texto tipo `O.PAGO "X" 0001-...`
  D Observaciones  → contexto (partner, ref de asiento, pago)
  E Debe / F Haber → `debit` / `credit`

Solo lectura: no escribe nada en Odoo.

Nota: el XLSX se arma vía CSV + LibreOffice headless (sin openpyxl en el entorno).
"""

from __future__ import annotations

import argparse
import csv
import subprocess
import sys
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

DEC2 = Decimal("0.01")


def _d2(x: Any) -> Decimal:
    if x is None or x is False:
        return Decimal("0")
    try:
        return Decimal(str(x))
    except Exception:
        return Decimal("0")


def _fmt_money_cell(x: Decimal) -> str:
    q = x.quantize(DEC2, rounding=ROUND_HALF_UP)
    s = f"{q:.2f}"
    return "0" if s in ("0.00", "-0.00") else s.rstrip("0").rstrip(".")


def _ddmmyyyy(iso_date: str) -> str:
    dt = datetime.strptime(iso_date, "%Y-%m-%d").date()
    return dt.strftime("%d/%m/%Y")


def odoo_connect(cfg: dict) -> tuple[Any, int]:
    url = cfg["url"].rstrip("/")
    common = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/common", allow_none=True)
    uid = common.authenticate(cfg["db"], cfg["username"], cfg["password"], {})
    if not uid:
        raise SystemExit("Autenticación Odoo fallida (ODOO_CONFIG_MASTER_DEV).")
    return xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/object", allow_none=True), int(uid)


def _convert_csv_to_xlsx(csv_path: Path, xlsx_out: Path) -> Path:
    xlsx_out.parent.mkdir(parents=True, exist_ok=True)
    tmp_dir = xlsx_out.parent
    cmd = [
        "libreoffice",
        "--headless",
        "--nologo",
        "--nolockcheck",
        "--nodefault",
        "--norestore",
        "--convert-to",
        "xlsx",
        "--outdir",
        str(tmp_dir),
        str(csv_path),
    ]
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        raise SystemExit(
            "Fallo la conversión a XLSX con LibreOffice.\n"
            f"STDOUT:\n{r.stdout}\nSTDERR:\n{r.stderr}\n"
        )
    produced = tmp_dir / (csv_path.stem + ".xlsx")
    if not produced.exists():
        raise SystemExit(f"LibreOffice no generó el archivo esperado: {produced}")
    if produced.resolve() != xlsx_out.resolve():
        produced.replace(xlsx_out)
    return xlsx_out


def _iibb_tax_ids(models: Any, db: str, uid: int, pwd: str) -> list[int]:
    # OR: l10n_ar_tax_type ilike iibb, name ilike IIBB, name ilike SIRCAR
    dom: list[Any] = ["|", "|", ("l10n_ar_tax_type", "ilike", "iibb"), ("name", "ilike", "IIBB"), ("name", "ilike", "SIRCAR")]
    ids: list[int] = models.execute_kw(db, uid, pwd, "account.tax", "search", [dom])
    return [int(x) for x in ids]


def _sum_debe_haber(
    models: Any,
    db: str,
    uid: int,
    pwd: str,
    *,
    tax_ids: list[int],
    hasta_excl: str,
) -> tuple[Decimal, Decimal]:
    dom: list[Any] = [
        ("date", "<", hasta_excl),
        ("parent_state", "=", "posted"),
        ("tax_line_id", "in", tax_ids),
    ]
    try:
        groups = models.execute_kw(
            db,
            uid,
            pwd,
            "account.move.line",
            "read_group",
            [dom, ["debit:sum", "credit:sum"], []],
            {"lazy": False},
        )
        if groups:
            g0 = groups[0]
            return _d2(g0.get("debit", 0)), _d2(g0.get("credit", 0))
    except Exception:
        pass
    ids = models.execute_kw(db, uid, pwd, "account.move.line", "search", [dom])
    if not ids:
        return Decimal("0"), Decimal("0")
    total_d = Decimal("0")
    total_c = Decimal("0")
    step = 2000
    for i in range(0, len(ids), step):
        chunk = ids[i : i + step]
        rows = models.execute_kw(
            db,
            uid,
            pwd,
            "account.move.line",
            "read",
            [chunk],
            {"fields": ["debit", "credit"]},
        )
        for r in rows:
            total_d += _d2(r.get("debit"))
            total_c += _d2(r.get("credit"))
    return total_d, total_c


def generar(desde: str, hasta: str, *, out_xlsx: Path) -> Path:
    cfg = ODOO_CONFIG_MASTER_DEV
    models, uid = odoo_connect(cfg)
    db, pwd = cfg["db"], cfg["password"]

    tax_ids = _iibb_tax_ids(models, db, uid, pwd)
    if not tax_ids:
        raise SystemExit("No se encontraron impuestos IIBB/SIRCAR (por nombre o l10n_ar_tax_type).")

    line_dom: list[Any] = [
        ("date", ">=", desde),
        ("date", "<=", hasta),
        ("parent_state", "=", "posted"),
        ("tax_line_id", "in", tax_ids),
    ]
    line_ids: list[int] = models.execute_kw(
        db,
        uid,
        pwd,
        "account.move.line",
        "search",
        [line_dom],
        {"order": "date asc, id asc"},
    )
    if not line_ids:
        raise SystemExit("No hay líneas de retención IIBB/SIRCAR en el rango dado.")

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
                "debit",
                "credit",
                "name",
                "move_id",
                "account_id",
                "partner_id",
                "payment_id",
            ]
        },
    )

    move_ids = sorted({l["move_id"][0] for l in lines if l.get("move_id")})
    moves: dict[int, dict] = {}
    if move_ids:
        mr = models.execute_kw(
            db,
            uid,
            pwd,
            "account.move",
            "read",
            [move_ids],
            {"fields": ["id", "name", "ref", "journal_id"]},
        )
        moves = {int(r["id"]): r for r in mr}

    journal_ids = sorted({m["journal_id"][0] for m in moves.values() if m.get("journal_id")})
    journals: dict[int, dict] = {}
    if journal_ids:
        jr = models.execute_kw(
            db,
            uid,
            pwd,
            "account.journal",
            "read",
            [journal_ids],
            {"fields": ["id", "code", "name", "type"]},
        )
        journals = {int(r["id"]): r for r in jr}

    pay_ids = sorted({l["payment_id"][0] for l in lines if l.get("payment_id")})
    payments: dict[int, dict] = {}
    if pay_ids:
        pr = models.execute_kw(
            db,
            uid,
            pwd,
            "account.payment",
            "read",
            [pay_ids],
            {"fields": ["id", "name"]},
        )
        payments = {int(r["id"]): r for r in pr}

    partner_ids = sorted({l["partner_id"][0] for l in lines if l.get("partner_id")})
    partners: dict[int, dict] = {}
    if partner_ids:
        pr = models.execute_kw(
            db,
            uid,
            pwd,
            "res.partner",
            "read",
            [partner_ids],
            {"fields": ["id", "name"]},
        )
        partners = {int(r["id"]): r for r in pr}

    acc_ids = sorted({l["account_id"][0] for l in lines if l.get("account_id")})
    accounts: dict[int, dict] = {}
    if acc_ids:
        ar = models.execute_kw(
            db,
            uid,
            pwd,
            "account.account",
            "read",
            [acc_ids],
            {"fields": ["id", "code", "name"]},
        )
        accounts = {int(r["id"]): r for r in ar}

    d0, c0 = _sum_debe_haber(models, db, uid, pwd, tax_ids=tax_ids, hasta_excl=desde)

    out_dir = out_xlsx.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    tmp_csv = out_dir / (out_xlsx.stem + ".csv")

    headers = ["Fecha", "Transaccion", "Detalle", "Observaciones", "Debe", "Haber"]
    rows: list[list[str]] = []
    rows.append([""] * 6)
    rows.append(["Listado Mayor de Cuentas"] + [""] * 5)
    rows.append([""] * 6)
    rows.append(headers)
    rows.append(["", "", "S.Ini.", "", _fmt_money_cell(d0), _fmt_money_cell(c0)])

    for l in lines:
        pid = l["payment_id"][0] if l.get("payment_id") else None
        pay = payments.get(int(pid)) if pid else {}

        mid = l["move_id"][0] if l.get("move_id") else None
        move = moves.get(int(mid)) if mid else {}
        jid = move.get("journal_id")[0] if move.get("journal_id") else None
        journal = journals.get(int(jid)) if jid else {}
        jcode = (journal.get("code") or "").strip()
        mname = (move.get("name") or "").strip()

        line_label = (l.get("name") or "").strip()
        acc = accounts.get(int(l["account_id"][0])) if l.get("account_id") else {}
        acc_disp = ""
        if acc:
            acc_disp = f"{acc.get('code') or ''} {acc.get('name') or ''}".strip()

        detalle_parts: list[str] = []
        if line_label:
            detalle_parts.append(line_label)
        if pay.get("name") and mname:
            detalle_parts.append(f'O.PAGO "{jcode or "X"}" {mname}'.strip())
        elif mname:
            detalle_parts.append(mname)
        if not detalle_parts and acc_disp:
            detalle_parts.append(acc_disp)
        detalle = " | ".join(detalle_parts) if detalle_parts else (acc_disp or str(l.get("id")))

        partner = partners.get(int(l["partner_id"][0])) if l.get("partner_id") else {}
        obs_bits: list[str] = []
        if partner.get("name"):
            obs_bits.append(str(partner["name"]))
        if move.get("ref"):
            obs_bits.append(str(move["ref"]))
        if pay.get("name"):
            obs_bits.append(f"Pago {pay['name']}")
        obs_bits.append("Retención IIBB/SIRCAR")
        observaciones = " — ".join(obs_bits)

        d = _d2(l.get("debit"))
        c = _d2(l.get("credit"))
        rows.append(
            [
                _ddmmyyyy(l["date"]),
                str(int(l["id"])),
                detalle,
                observaciones,
                _fmt_money_cell(d) if d != Decimal("0") else "0",
                _fmt_money_cell(c) if c != Decimal("0") else "0",
            ]
        )

    with tmp_csv.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f, delimiter=",", lineterminator="\n")
        w.writerows(rows)

    return _convert_csv_to_xlsx(tmp_csv, out_xlsx)


def _iso(s: str) -> str:
    datetime.strptime(s, "%Y-%m-%d")
    return s


def _default_range_mes_actual() -> tuple[str, str]:
    hoy = date.today()
    desde = hoy.replace(day=1)
    if desde.month == 12:
        hasta = desde.replace(year=desde.year + 1, month=1, day=1)
    else:
        hasta = desde.replace(month=desde.month + 1, day=1)
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
        default=Path(__file__).resolve().parent.parent / "out" / "RET_IIBB_odoo.xlsx",
        help="Ruta del XLSX a generar",
    )
    args = ap.parse_args(list(argv) if argv is not None else None)
    out = generar(args.desde, args.hasta, out_xlsx=args.out)
    print(f"OK: {out}")
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())

