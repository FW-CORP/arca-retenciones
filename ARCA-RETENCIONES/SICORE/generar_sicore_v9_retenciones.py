#!/usr/bin/env python3
"""
Genera TXT estándar SICORE v9.0 para importación de RETENCIONES (AFIP/ARCA).

Criterio estudio (2025): registro de **159** posiciones + CRLF (extiende el diseño base SICORE v9).
- N° comprobante (16): solo **dígitos**; 4 = punto de venta + 12 = número (sin guiones).
  Textos tipo FACOM… no son válidos fiscalmente: se derivan dígitos o se rellena con ceros.
- Código impuesto **0217** (4) + código régimen **3** dígitos (p. ej. `078`) desde `l10n_ar_code`.
- **Código de operación** (Tabla C ARCA): **1** = Retención (no usar `0`; no figura en la tabla).
- Montos sin separadores; últimos 2 dígitos = centavos.
- Tras importe retención: bloque de **13 ceros** (LPAD numérico; reemplaza % exclusión + fecha boletín + ajuste) para que **tipo doc** (`80`) inicie en columna **107** y el CUIT no se desplace.
- Documento retenido: **22** dígitos (CUIT 11 + relleno 11 ceros).
- Cola pos. 131-159: último campo **1 espacio** (relleno final); no dos ceros.
- **No confundir**: el **código de comprobante** (Tabla A) son las pos. 1–2 (`01` Factura, `06` OP, etc.). El **`80`** es el **tipo de documento del retenido** (Tabla F: CUIT), pos. **107–108**.
- **Código condición** (Tabla D, pos. 77–78, 2 caracteres): **`01`** = Inscripto; no es un solo dígito `1`.

Solo lectura: Odoo `master_dev` (XML-RPC).
"""

from __future__ import annotations

import argparse
import re
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

# Longitud total registro (sin CRLF)
LREG = 159

# Orden y anchos para desglose / validación (debe coincidir con `_reg_sicore_line_160`)
SICORE_V9_160_CAMPOS: tuple[tuple[str, int], ...] = (
    ("cod_comprobante", 2),
    ("fecha_comprobante", 10),
    ("nro_comprobante", 16),
    ("importe_comprobante", 16),
    ("cod_impuesto", 4),
    ("cod_regimen", 3),
    ("cod_operacion", 1),
    ("base", 14),
    ("fecha_retencion", 10),
    ("cod_condicion", 2),
    ("importe_retencion", 14),  # Campo 11 (079-092) según grilla estudio
    ("excedente_otros", 14),  # Campo 12 (093-106) según grilla estudio
    ("tipo_doc", 2),
    ("nro_cuit", 11),  # Campo 14 (109-119)
    ("porc_exclusion", 11),  # Campo 15 (120-130): ceros si no aplica
    ("fecha_pub_certificado", 10),
    ("tipo_regimen_especial", 2),
    ("importe_base_exclusion", 14),
    ("tipo_cuenta", 2),
    ("relleno_final", 1),
)


def validar_posiciones_clave_sicore(path: Path) -> tuple[int, list[tuple[int, list[str]]]]:
    """
    Comprueba en cada línea (columnas 1-based, criterio Bloc de notas):
    - Col 67: inicio fecha retención (dd/mm/aaaa)
    - Col 107: inicio tipo doc (`80`)
    - Col 109: primer dígito CUIT
    """
    raw = path.read_text(encoding="ascii", errors="replace")
    col_fecha, col_tipo, col_cuit = 67, 107, 109
    i_f, i_t, i_c = col_fecha - 1, col_tipo - 1, col_cuit - 1
    failures: list[tuple[int, list[str]]] = []
    n_ok = 0
    for lineno, line in enumerate(raw.splitlines(), 1):
        s = line.rstrip("\r\n")
        if not s:
            continue
        errs: list[str] = []
        if len(s) != LREG:
            errs.append(f"largo {len(s)} != {LREG}")
        if len(s) > i_f + 9:
            chunk = s[i_f : i_f + 10]
            if not (len(chunk) == 10 and chunk[2] == "/" and chunk[5] == "/"):
                errs.append(f"col {col_fecha}: esperaba dd/mm/aaaa, got {chunk!r}")
        else:
            errs.append(f"col {col_fecha}: línea corta")
        if len(s) > i_t + 1:
            if s[i_t : i_t + 2] != "80":
                errs.append(f"col {col_tipo}: esperaba '80', got {s[i_t:i_t+2]!r}")
        else:
            errs.append(f"col {col_tipo}: línea corta")
        if len(s) > i_c:
            if not s[i_c].isdigit():
                errs.append(f"col {col_cuit}: esperaba dígito, got {s[i_c]!r}")
        else:
            errs.append(f"col {col_cuit}: línea corta")
        if errs:
            failures.append((lineno, errs))
        else:
            n_ok += 1
    return n_ok, failures


def desglosar_registro_sicore_160(line: str) -> list[tuple[str, str]]:
    """Parte una línea (sin CRLF) en campos según `SICORE_V9_160_CAMPOS` (LREG posiciones)."""
    s = line.rstrip("\r\n")
    if len(s) != LREG:
        raise ValueError(f"Largo {len(s)} != {LREG}")
    out: list[tuple[str, str]] = []
    i = 0
    for name, w in SICORE_V9_160_CAMPOS:
        out.append((name, s[i : i + w]))
        i += w
    return out


def odoo_connect(cfg: dict) -> tuple[Any, int]:
    url = cfg["url"].rstrip("/")
    common = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/common", allow_none=True)
    uid = common.authenticate(cfg["db"], cfg["username"], cfg["password"], {})
    if not uid:
        raise SystemExit("Autenticación Odoo fallida (ODOO_CONFIG_MASTER_DEV).")
    return xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/object", allow_none=True), int(uid)


def _d(x: Any) -> Decimal:
    if x is None or x is False:
        return Decimal("0")
    try:
        return Decimal(str(x))
    except Exception:
        return Decimal("0")


def _abs_money(x: Any) -> Decimal:
    return abs(_d(x)).quantize(DEC2, rounding=ROUND_HALF_UP)


def _digits_only(s: str) -> str:
    return re.sub(r"\D+", "", s or "")


def _many2one_id(val: Any) -> int | None:
    if val is None or val is False:
        return None
    if isinstance(val, (list, tuple)) and val:
        return int(val[0])
    if isinstance(val, int):
        return int(val)
    return None


def _ddmmyyyy(iso_date: str) -> str:
    return datetime.strptime(iso_date, "%Y-%m-%d").date().strftime("%d/%m/%Y")


def _fmt_centavos_sin_separador(value: Decimal, width: int) -> str:
    cents = (value * 100).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
    try:
        n = int(cents)
    except Exception:
        n = 0
    if n < 0:
        n = -n
    s = str(n).zfill(width)
    return s[-width:]


def _nro_comprobante_16_fiscal(raw: str) -> str:
    """
    16 dígitos (layout estudio):
    - Si viene `PV-NUM` (o `PV/NUM`) con NUM hasta 8 dígitos:
        0000 + PV(4) + NUM(8)
        Ej: 2526-30094 → 0000252600030094
    - Si NUM excede 8 dígitos o el patrón no es claro, se cae a una normalización conservadora
      (solo dígitos y padding) para siempre devolver 16 caracteres.
    Prefijos alfanuméricos (FACOM, etc.) se ignoran; solo dígitos útiles.
    """
    s = (raw or "").strip()
    if not s:
        return "0" * 16

    m = re.match(r"^\s*(\d{1,9})\s*[-/]\s*(\d{1,12})\s*$", s)
    if m:
        pv = int(m.group(1))
        num = int(m.group(2))
        pv = max(0, min(pv, 9999))
        # Si el número entra en 8 dígitos, respetar 0000 + PV(4) + NUM(8)
        if num <= 99999999:
            return f"0000{pv:04d}{num:08d}"
        num = max(0, min(num, 999999999999))
        return f"{pv:04d}{num:012d}"

    dig = _digits_only(s)
    if not dig:
        return "0" * 16

    # Sin guión: tomar últimos dígitos como número fiscal (evitar letras FACOM…)
    if len(dig) >= 16:
        return dig[-16:]
    if len(dig) >= 10:
        pv = int(dig[:4])
        rest = dig[4:]
        num = int(rest) if rest else 0
        return f"{pv:04d}{num:012d}"
    return dig.zfill(16)[-16:]


def _cod_regimen_3_desde_tax_code(reg_raw: str) -> str:
    """3 dígitos (078), no 0078."""
    d = _digits_only(str(reg_raw or "").strip())
    if not d:
        return "000"
    if len(d) >= 3:
        return d[-3:].zfill(3)
    return d.zfill(3)


def _fmt_cuit_11(cuit11: str) -> str:
    """CUIT 11 dígitos, solo números, con cero a la izquierda si falta."""
    d = _digits_only(cuit11)
    if len(d) > 11:
        d = d[-11:]
    elif len(d) < 11:
        d = d.zfill(11)
    return d[:11]


def _reg_sicore_line_160(
    *,
    cod_comprobante_2: str,
    fecha_comprobante_10: str,
    nro_comprobante_16: str,
    importe_comprobante_16: str,
    cod_impuesto_4: str,
    cod_regimen_3: str,
    cod_operacion_1: str,
    base_14: str,
    fecha_retencion_10: str,
    cod_condicion_2: str,
    importe_retencion_14: str,
    excedente_otros_14: str,
    tipo_doc_2: str,
    nro_cuit_11: str,
    porc_exclusion_11: str,
    # 131-158 + 159
    fecha_pub_certificado_10: str,
    tipo_regimen_especial_2: str,
    importe_base_exclusion_14: str,
    tipo_cuenta_2: str,
    relleno_final_1: str,
) -> str:
    core = (
        cod_comprobante_2
        + fecha_comprobante_10
        + nro_comprobante_16
        + importe_comprobante_16
        + cod_impuesto_4
        + cod_regimen_3
        + cod_operacion_1
        + base_14
        + fecha_retencion_10
        + cod_condicion_2
        + importe_retencion_14
        + excedente_otros_14
        + tipo_doc_2
        + nro_cuit_11
        + porc_exclusion_11
    )
    tail = (
        fecha_pub_certificado_10
        + tipo_regimen_especial_2
        + importe_base_exclusion_14
        + tipo_cuenta_2
        + relleno_final_1
    )
    out = core + tail
    if len(out) != LREG:
        raise ValueError(f"Largo registro {len(out)} != {LREG}")
    return out


def generar(
    desde: str,
    hasta: str,
    *,
    out_path: Path,
    modo_comprobante: str = "auto",
    fecha_comprobante: str = "auto",
    codigo_comprobante_manual: str | None = None,
    codigo_comprobante_sin_factura: str = "06",
    codigo_impuesto_ganancias: str = "0217",
    codigo_operacion: str = "0",
    codigo_condicion: str = "01",
    tipo_doc: str = "80",
) -> Path:
    cfg = ODOO_CONFIG_MASTER_DEV
    models, uid = odoo_connect(cfg)
    db, pwd = cfg["db"], cfg["password"]

    earn_tax_ids: list[int] = models.execute_kw(
        db,
        uid,
        pwd,
        "account.tax",
        "search",
        [[("l10n_ar_tax_type", "=", "earnings")]],
    )
    if not earn_tax_ids:
        raise SystemExit("No hay impuestos con l10n_ar_tax_type='earnings' en la base.")
    earn_set = set(int(x) for x in earn_tax_ids)

    pay_domain = [
        ("date", ">=", desde),
        ("date", "<=", hasta),
        ("partner_type", "=", "supplier"),
        ("l10n_ar_withholding_ids", "!=", False),
        ("state", "in", ["paid", "in_process", "posted"]),
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
        {
            "fields": [
                "id",
                "name",
                "date",
                "partner_id",
                "amount",
                "l10n_ar_withholding_ids",
                "reconciled_bill_ids",
            ]
        },
    )

    all_bill_ids: set[int] = set()
    for p in pays:
        for t in p.get("reconciled_bill_ids") or []:
            bid = _many2one_id(t)
            if bid is not None:
                all_bill_ids.add(bid)

    bills_map: dict[int, dict] = {}
    if all_bill_ids:
        br = models.execute_kw(
            db,
            uid,
            pwd,
            "account.move",
            "read",
            [sorted(all_bill_ids)],
            {
                "fields": [
                    "id",
                    "name",
                    "invoice_date",
                    "amount_total",
                    "move_type",
                    "l10n_latam_document_number",
                ]
            },
        )
        bills_map = {int(r["id"]): r for r in br}

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
            {"fields": ["id", "vat", "l10n_ar_vat"]},
        )
        partners = {int(r["id"]): r for r in pr}

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
        tr = models.execute_kw(
            db,
            uid,
            pwd,
            "account.tax",
            "read",
            [tax_ids],
            {"fields": ["id", "name", "l10n_ar_tax_type", "l10n_ar_code"]},
        )
        taxes = {int(r["id"]): r for r in tr}

    lines_by_payment: dict[int, list[dict]] = {}
    for l in lines:
        if not l.get("payment_id") or not l.get("tax_line_id"):
            continue
        tid = int(l["tax_line_id"][0])
        tax = taxes.get(tid) or {}
        if tid not in earn_set or tax.get("l10n_ar_tax_type") != "earnings":
            continue
        pid = int(l["payment_id"][0])
        lines_by_payment.setdefault(pid, []).append(l)

    # Cola fija 131-159 (valores por defecto estudio / grilla importación)
    fecha_pub_def = "00/00/0000"
    tipo_reg_esp_def = "00"
    base_excl_def = "0" * 14
    tipo_cta_def = "00"
    relleno_final_def = " "

    out_lines: list[str] = []
    for p in pays:
        pid = int(p["id"])
        wlines = lines_by_payment.get(pid) or []
        if not wlines:
            continue

        fecha_ret = _ddmmyyyy(p["date"])
        partner_id = int(p["partner_id"][0]) if p.get("partner_id") else 0
        partner = partners.get(partner_id, {})
        cuit_raw = str(partner.get("l10n_ar_vat") or partner.get("vat") or "")
        cuit11 = _digits_only(cuit_raw)
        if len(cuit11) > 11:
            cuit11 = cuit11[-11:]
        nro_cuit_11 = _fmt_cuit_11(cuit11)

        bill_tuples = p.get("reconciled_bill_ids") or []
        bills_ord: list[dict] = []
        for bt in bill_tuples:
            bid = _many2one_id(bt)
            if bid is None:
                continue
            b = bills_map.get(bid)
            if b:
                bills_ord.append(b)
        bills_ord.sort(key=lambda r: (r.get("invoice_date") or "", int(r["id"])))

        use_bill = bool(bills_ord) and modo_comprobante in ("auto", "factura")
        if modo_comprobante == "orden_pago":
            use_bill = False
        if modo_comprobante == "factura" and not bills_ord:
            continue

        # Fecha de comprobante: auto=según fuente, pago=siempre fecha del pago, factura=siempre fecha factura (si existe)
        if fecha_comprobante not in ("auto", "pago", "factura"):
            raise SystemExit(f"--fecha-comprobante inválida: {fecha_comprobante!r}")

        if codigo_comprobante_manual:
            cod_comp2 = str(codigo_comprobante_manual).zfill(2)[:2]
            fecha_comp = fecha_ret
            raw_nro = str(p.get("name") or "")
            nro_comp_16 = _nro_comprobante_16_fiscal(raw_nro)
            imp_comp = _fmt_centavos_sin_separador(_abs_money(p.get("amount")), 16)
        elif use_bill:
            bill = bills_ord[0]
            cod_comp2 = "01"
            idate = bill.get("invoice_date") or p["date"]
            fecha_comp = _ddmmyyyy(str(idate))
            raw_nro = (
                str(bill.get("l10n_latam_document_number") or "").strip()
                or str(bill.get("name") or "").strip()
            )
            nro_comp_16 = _nro_comprobante_16_fiscal(raw_nro)
            imp_comp = _fmt_centavos_sin_separador(_abs_money(bill.get("amount_total")), 16)
        else:
            cod_comp2 = str(codigo_comprobante_sin_factura).zfill(2)[:2]
            fecha_comp = fecha_ret
            raw_nro = str(p.get("name") or "")
            nro_comp_16 = _nro_comprobante_16_fiscal(raw_nro)
            imp_comp = _fmt_centavos_sin_separador(_abs_money(p.get("amount")), 16)

        # Override de fecha de comprobante según configuración
        if fecha_comprobante == "pago":
            fecha_comp = fecha_ret
        elif fecha_comprobante == "factura":
            if use_bill:
                # ya está seteada desde factura
                pass
            else:
                # si no hay factura, caer a fecha del pago
                fecha_comp = fecha_ret

        for l in sorted(wlines, key=lambda r: int(r["id"])):
            base = _fmt_centavos_sin_separador(_abs_money(l.get("tax_base_amount")), 14)
            imp_ret = _fmt_centavos_sin_separador(
                _abs_money(l.get("credit") or l.get("balance") or 0), 14
            )
            if imp_ret == "0" * 14:
                imp_ret = _fmt_centavos_sin_separador(_abs_money(l.get("balance")), 14)

            tax = taxes.get(int(l["tax_line_id"][0])) if l.get("tax_line_id") else {}
            cod_reg_3 = _cod_regimen_3_desde_tax_code(str(tax.get("l10n_ar_code") or ""))

            line = _reg_sicore_line_160(
                cod_comprobante_2=cod_comp2,
                fecha_comprobante_10=fecha_comp,
                nro_comprobante_16=nro_comp_16,
                importe_comprobante_16=imp_comp,
                cod_impuesto_4=str(codigo_impuesto_ganancias).zfill(4)[:4],
                cod_regimen_3=cod_reg_3,
                cod_operacion_1=str(codigo_operacion)[:1],
                base_14=base,
                fecha_retencion_10=fecha_ret,
                cod_condicion_2=str(codigo_condicion).zfill(2)[:2],
                importe_retencion_14=imp_ret,
                excedente_otros_14="0" * 14,
                tipo_doc_2=str(tipo_doc).zfill(2)[:2],
                nro_cuit_11=nro_cuit_11,
                porc_exclusion_11="0" * 11,
                fecha_pub_certificado_10=fecha_pub_def,
                tipo_regimen_especial_2=tipo_reg_esp_def,
                importe_base_exclusion_14=base_excl_def,
                tipo_cuenta_2=tipo_cta_def,
                relleno_final_1=relleno_final_def,
            )
            out_lines.append(line)

    if not out_lines:
        raise SystemExit("No se encontraron retenciones de Ganancias (earnings) en el rango.")

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
    ap.add_argument(
        "--desglosar",
        type=Path,
        metavar="ARCHIVO",
        help="Lee la primera línea del TXT (LREG pos), imprime campos posición a posición y termina (sin Odoo).",
    )
    ap.add_argument("--desde", type=_iso, default=d_desde)
    ap.add_argument("--hasta", type=_iso, default=d_hasta)
    ap.add_argument(
        "--out",
        type=Path,
        default=Path(__file__).resolve().parent / "out" / "SICORE_V9_RETENCIONES_GANANCIAS.TXT",
    )
    ap.add_argument(
        "--modo-comprobante",
        choices=("auto", "factura", "orden_pago"),
        default="auto",
        help="auto=01+factura si hay conciliación; factura=solo con factura; orden_pago=06+pago",
    )
    ap.add_argument(
        "--fecha-comprobante",
        choices=("auto", "pago", "factura"),
        default="auto",
        help="Qué fecha usar en campo 2: auto=según fuente; pago=siempre fecha de retención/pago; factura=fecha factura si existe.",
    )
    ap.add_argument("--codigo-comprobante", default=None, help="Forzar 01/06 (anula modo)")
    ap.add_argument("--codigo-comprobante-sin-factura", default="06")
    ap.add_argument("--codigo-impuesto", default="0217")
    ap.add_argument(
        "--codigo-operacion",
        default="0",
        help="Según grilla estudio: generalmente 0. (Si tu configuración SICORE usa Tabla C estándar: 1=Retención).",
    )
    ap.add_argument(
        "--codigo-condicion",
        default="01",
        help="Tabla D (v9), 2 dígitos, pos.77-78: 01=Inscripto, 00=Ninguna (según régimen).",
    )
    ap.add_argument(
        "--tipo-doc",
        default="80",
        help="Tabla F: tipo doc. del retenido (pos.107-108). 80=CUIT. No es el código de comprobante 01/06.",
    )
    ap.add_argument(
        "--validar-posiciones",
        type=Path,
        metavar="ARCHIVO",
        help="Valida cols 67 (fecha ret.), 107 (80), 109 (CUIT) en todas las líneas; termina sin Odoo.",
    )
    args = ap.parse_args(list(argv) if argv is not None else None)

    if args.validar_posiciones:
        path = args.validar_posiciones.expanduser()
        if not path.is_file():
            raise SystemExit(f"No existe el archivo: {path}")
        n_ok, failures = validar_posiciones_clave_sicore(path)
        if failures:
            print(f"FALLO: {len(failures)} línea(s), {n_ok} OK")
            for lineno, errs in failures[:20]:
                print(f"  Línea {lineno}: {errs}")
            if len(failures) > 20:
                print(f"  ... y {len(failures) - 20} más")
            return 1
        print(f"OK: {n_ok} línea(s) — col 67 fecha retención, col 107 '80', col 109 CUIT.")
        return 0

    if args.desglosar:
        path = args.desglosar.expanduser()
        if not path.is_file():
            raise SystemExit(f"No existe el archivo: {path}")
        raw = path.read_text(encoding="ascii", errors="replace")
        first = raw.splitlines()[0] if raw else ""
        try:
            campos = desglosar_registro_sicore_160(first)
        except ValueError as e:
            raise SystemExit(f"Primera línea inválida: {e}\n{first!r}") from e
        pos = 1
        for nombre, valor in campos:
            w = len(valor)
            print(f"{pos:3}-{pos + w - 1:3}  {nombre:22}  {valor!r}")
            pos += w
        return 0

    out = generar(
        args.desde,
        args.hasta,
        out_path=args.out,
        modo_comprobante=args.modo_comprobante,
        fecha_comprobante=args.fecha_comprobante,
        codigo_comprobante_manual=args.codigo_comprobante,
        codigo_comprobante_sin_factura=args.codigo_comprobante_sin_factura,
        codigo_impuesto_ganancias=args.codigo_impuesto,
        codigo_operacion=args.codigo_operacion,
        codigo_condicion=args.codigo_condicion,
        tipo_doc=args.tipo_doc,
    )
    print(f"OK: {out}")
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
