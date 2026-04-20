#!/usr/bin/env python3
"""
Genera una planilla `OP_odoo.xlsx` con el mismo esquema que `OP.xlsx`,
leyendo pagos a proveedores desde Odoo `master_dev` por XML-RPC.

Restricciones:
- SOLO LECTURA (no escribe nada en Odoo).

Columnas (según OP.xlsx):
 A) FECHA
 B) Nro. Orden de Pago (account.payment.name)
 C) Nombre Proveedor
 D) Importe Pago (account.payment.amount)
 E) (ignorar)
 F) (ignorar)
 G) Cheques
 H) Transferencias
 I) Cheques Terceros
 J) Retenciones (total)
 K) Ret. Ing. Brutos (subtotal IIBB/SIRCAR)
 L) Efectivo

Odoo vs planilla:
- Un diario puede tener **Tipo = Efectivo** en Odoo y aun así ser **“Cheques de Terceros”**
  (ej. código `CHQS`): para la columna **I** (cheques de terceros), no usar la **L** (efectivo).
- La columna **L** corresponde al diario **Efectivo** real (ej. código `EFVO`).

Nota técnica:
En este entorno no hay `pip/openpyxl`, así que el XLSX se genera creando un CSV
y convirtiéndolo a XLSX con LibreOffice en modo headless.
"""

from __future__ import annotations

import argparse
import csv
import re
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


def _abs_money(x: Any) -> Decimal:
    return abs(_d2(x)).quantize(DEC2, rounding=ROUND_HALF_UP)


def _fmt_money(x: Decimal) -> str:
    # CSV del template OP suele estar en enteros o decimales; usamos punto.
    return f"{x:.2f}".rstrip("0").rstrip(".") if "." in f"{x:.2f}" else f"{x:.2f}"


def _ddmmyyyy(iso_date: str) -> str:
    dt = datetime.strptime(iso_date, "%Y-%m-%d").date()
    return dt.strftime("%d/%m/%Y")


def _is_iibb_tax(tax: dict) -> bool:
    nm = (tax.get("name") or "").upper()
    tt = (tax.get("l10n_ar_tax_type") or "").lower()
    return tt.startswith("iibb") or ("SIRCAR" in nm) or ("IIBB" in nm)


def odoo_connect(cfg: dict) -> tuple[Any, int]:
    url = cfg["url"].rstrip("/")
    common = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/common", allow_none=True)
    uid = common.authenticate(cfg["db"], cfg["username"], cfg["password"], {})
    if not uid:
        raise SystemExit("Autenticación Odoo fallida (revisá ODOO_CONFIG_MASTER_DEV).")
    return xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/object", allow_none=True), int(uid)


def _payment_prefix(name: str) -> str:
    """Prefijo del número de pago antes de la primera barra (ej. PGAL1/26-27/0002 -> PGAL1)."""
    n = (name or "").strip().upper()
    if not n:
        return ""
    return n.split("/", 1)[0]


def _bucket_from_payment_name(name: str) -> str | None:
    n = (name or "").upper().strip()
    if not n:
        return None
    pref = _payment_prefix(name)
    # PCHQS antes que PCHQ (si no, “cheques de terceros” cae en columna G por error).
    if pref.startswith("PCHQS") or pref == "CHQS":
        return "cheque_terceros"
    if pref.startswith("PEFVO") or pref.startswith("EFVO"):
        return "efectivo"
    if pref.startswith("PCHQ") or ("CHQ" in n and "TERC" not in n):
        return "cheque"
    if pref.startswith("PBNA") or pref.startswith("PGAL") or "TRANS" in n or "TRF" in n or "BNA" in n:
        return "transfer"
    if "TERC" in n:
        return "cheque_terceros"
    if "CAJ" in n or "EFEC" in n:
        return "efectivo"
    return None


def _bucket_from_journal_record(j: dict) -> str | None:
    """Usa código corto y nombre del diario (más fiable que `journal.type` para la planilla)."""
    code = (j.get("code") or "").upper().strip()
    name = (j.get("name") or "").upper().strip()
    # Nakel: CHQS = Cheques de Terceros (aunque en Odoo el tipo sea “cash”).
    if code == "CHQS" or ("TERCER" in name and ("CHEQ" in name or "CHQ" in name)):
        return "cheque_terceros"
    if code == "EFVO" or (name == "EFECTIVO" or name.startswith("EFECTIVO ")):
        return "efectivo"
    if "TERCER" in name or "3RO" in name:
        return "cheque_terceros"
    if "CHEQ" in name or "CHEQUE" in name:
        return "cheque"
    if "TRANS" in name or "TRANSFER" in name:
        return "transfer"
    if "CAJA" in name or (name == "EFECTIVO"):
        return "efectivo"
    if "BANCO" in name or "BNA" in name:
        return "transfer"
    jtype = (j.get("type") or "").lower()
    if jtype == "bank":
        return "transfer"
    return None


def _bucket_payment(payment: dict, journals: dict[int, dict]) -> str:
    j: dict | None = None
    jid_tuple = payment.get("journal_id")
    if jid_tuple and isinstance(jid_tuple, (list, tuple)) and jid_tuple:
        j = journals.get(int(jid_tuple[0])) or {}

    # 1) Diario (código CHQS / EFVO resuelve el caso “tipo Efectivo” vs columnas I/L).
    if j:
        b = _bucket_from_journal_record(j)
        if b:
            return b
    # 2) Número de pago (secuencia dedicada, ej. PCHQS/…).
    b = _bucket_from_payment_name(payment.get("name") or "")
    if b:
        return b
    # 3) Fallback: banco vs resto (no asumir “efectivo” solo por journal.type=cash).
    if j and str(j.get("type") or "").lower() == "bank":
        return "transfer"
    return "transfer"


def _convert_csv_to_xlsx(csv_path: Path, xlsx_out: Path) -> Path:
    xlsx_out.parent.mkdir(parents=True, exist_ok=True)
    # LibreOffice escribe el output en el directorio indicado (con el mismo basename).
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


def generar(
    desde: str,
    hasta: str,
    *,
    out_xlsx: Path,
) -> Path:
    cfg = ODOO_CONFIG_MASTER_DEV
    models, uid = odoo_connect(cfg)
    db, pwd = cfg["db"], cfg["password"]

    # En master_dev los estados de pago pueden variar según el flujo; para no perder
    # registros del Excel de OP, filtramos por fechas + proveedor y excluimos cancelados.
    pay_domain = [
        ("date", ">=", desde),
        ("date", "<=", hasta),
        ("partner_type", "=", "supplier"),
        ("state", "!=", "cancelled"),
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
        raise SystemExit("No se encontraron pagos a proveedores (posted) en el rango dado.")

    pays: list[dict] = models.execute_kw(
        db,
        uid,
        pwd,
        "account.payment",
        "read",
        [pay_ids],
        {
            "fields": [
                "id",
                "name",
                "date",
                "amount",
                "partner_id",
                "journal_id",
                "l10n_ar_withholding_ids",
            ]
        },
    )

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
            {"fields": ["id", "name"]},
        )
        partners = {int(r["id"]): r for r in pr}

    journal_ids = sorted({p["journal_id"][0] for p in pays if p.get("journal_id")})
    journals: dict[int, dict] = {}
    if journal_ids:
        jr = models.execute_kw(
            db,
            uid,
            pwd,
            "account.journal",
            "read",
            [journal_ids],
            {"fields": ["id", "name", "code", "type"]},
        )
        journals = {int(r["id"]): r for r in jr}

    line_ids = sorted(
        {lid for p in pays for lid in (p.get("l10n_ar_withholding_ids") or [])}
    )
    lines: list[dict] = []
    if line_ids:
        lines = models.execute_kw(
            db,
            uid,
            pwd,
            "account.move.line",
            "read",
            [line_ids],
            {
                "fields": [
                    "id",
                    "payment_id",
                    "credit",
                    "debit",
                    "balance",
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
            {"fields": ["id", "name", "l10n_ar_tax_type"]},
        )
        taxes = {int(r["id"]): r for r in tr}

    lines_by_payment: dict[int, list[dict]] = {}
    for l in lines:
        pid = l["payment_id"][0] if l.get("payment_id") else None
        if not pid:
            continue
        lines_by_payment.setdefault(int(pid), []).append(l)

    # CSV con 12 columnas. Replicamos “estructura” del OP.xlsx (títulos + headers).
    out_dir = out_xlsx.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    tmp_csv = out_dir / (out_xlsx.stem + ".csv")

    headers = [
        "Fecha",
        "Nro.Orden pago",
        "Nombre Proveedor",
        "Importe Pago",
        "Importe Ticket",
        "Trajetas",
        "Cheques",
        "Transferencias",
        "Cheques Terceros",
        "Retenciones",
        "Ret.Ing.Brutos",
        "Efectivo",
    ]

    # 1) filas “fijas” (como en OP.xlsx)
    rows: list[list[str]] = []
    rows.append([""] * 12)  # fila 1 vacía
    rows.append(["Listado de Ordenes de Pago"] + [""] * 11)  # fila 2
    rows.append([""] * 12)  # fila 3
    rows.append([""] * 12)  # fila 4
    rows.append(["Nakel S.A."] + [""] * 11)  # fila 5
    rows.append([""] * 12)  # fila 6
    rows.append(headers)  # fila 7

    for p in sorted(pays, key=lambda r: (r.get("date") or "", int(r["id"]))):
        pid = int(p["id"])
        partner_id = p["partner_id"][0] if p.get("partner_id") else None
        proveedor = (partners.get(int(partner_id)) or {}).get("name") if partner_id else ""

        monto_pago = _abs_money(p.get("amount"))
        bucket = _bucket_payment(p, journals=journals)

        cheque = Decimal("0.00")
        transferencia = Decimal("0.00")
        cheque_terceros = Decimal("0.00")
        efectivo = Decimal("0.00")

        if bucket == "cheque":
            cheque = monto_pago
        elif bucket == "transfer":
            transferencia = monto_pago
        elif bucket == "cheque_terceros":
            cheque_terceros = monto_pago
        elif bucket == "efectivo":
            efectivo = monto_pago
        else:
            transferencia = monto_pago

        # retenciones
        total_ret = Decimal("0.00")
        total_iibb = Decimal("0.00")
        for l in lines_by_payment.get(pid, []):
            imp = _abs_money(l.get("credit") or l.get("balance") or l.get("debit") or 0)
            total_ret += imp
            tax = taxes.get(l["tax_line_id"][0]) if l.get("tax_line_id") else None
            if tax and _is_iibb_tax(tax):
                total_iibb += imp

        rows.append(
            [
                _ddmmyyyy(p["date"]),  # A Fecha
                str(p.get("name") or ""),  # B Nro Orden de Pago
                str(proveedor or ""),  # C
                _fmt_money(monto_pago),  # D
                "",  # E ignorar
                "",  # F ignorar
                _fmt_money(cheque) if cheque != Decimal("0.00") else "",  # G
                _fmt_money(transferencia) if transferencia != Decimal("0.00") else "",  # H
                _fmt_money(cheque_terceros) if cheque_terceros != Decimal("0.00") else "",  # I
                _fmt_money(total_ret) if total_ret != Decimal("0.00") else "",  # J
                _fmt_money(total_iibb) if total_iibb != Decimal("0.00") else "",  # K
                _fmt_money(efectivo) if efectivo != Decimal("0.00") else "",  # L
            ]
        )

    with tmp_csv.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f, delimiter=",", lineterminator="\n")
        w.writerows(rows)

    out = _convert_csv_to_xlsx(tmp_csv, out_xlsx)
    return out


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
        default=Path(__file__).resolve().parent.parent / "out" / "OP_odoo.xlsx",
        help="Ruta del XLSX a generar",
    )
    args = ap.parse_args(list(argv) if argv is not None else None)

    out = generar(args.desde, args.hasta, out_xlsx=args.out)
    print(f"OK: {out}")
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())

