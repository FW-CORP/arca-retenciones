#!/usr/bin/env python3
"""
Genera certificados de retención en PDF (masivo) desde Odoo `master_dev`.

Estrategia (sin dependencias extra):
- Usa la plantilla Excel provista por el estudio (XLSM) y la convierte a XLSX con LibreOffice.
- Rellena celdas de la hoja `LOCAL` (Print_Area) con datos de Odoo.
- Exporta a PDF con LibreOffice headless.

Fuente Odoo:
- `account.move.line` con `name ilike 'CERT-'` y `tax_line_id` (línea de retención).
  El número de retención sale de `account.move.line.name` (ej. CERT-2026-086168).

Solo lectura: no escribe en Odoo.
"""

from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
import unicodedata
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Iterable

import xml.etree.ElementTree as ET
import xmlrpc.client


DEC2 = Decimal("0.01")
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def _d2(x: Any) -> Decimal:
    if x is None or x is False:
        return Decimal("0")
    try:
        return Decimal(str(x))
    except Exception:
        return Decimal("0")


def _abs_money(x: Any) -> Decimal:
    return abs(_d2(x)).quantize(DEC2, rounding=ROUND_HALF_UP)


def _excel_date_serial(d: date) -> int:
    # Excel (Windows) epoch para Calc/OOXML: 1899-12-30
    epoch = date(1899, 12, 30)
    return (d - epoch).days


def _digits_only(s: str) -> str:
    return re.sub(r"\D+", "", s or "")


def _cuit11_from_partner(row: dict) -> str:
    raw = row.get("l10n_ar_vat") or row.get("vat") or ""
    dig = _digits_only(str(raw))
    if len(dig) == 11:
        return dig
    if len(dig) > 11:
        return dig[-11:]
    return ""


def _num_to_words_es_pesos(amount: Decimal) -> str:
    n = amount.quantize(DEC2, rounding=ROUND_HALF_UP)
    entero = int(n)
    cent = int((n - Decimal(entero)) * 100)

    unidades = [
        "CERO",
        "UNO",
        "DOS",
        "TRES",
        "CUATRO",
        "CINCO",
        "SEIS",
        "SIETE",
        "OCHO",
        "NUEVE",
        "DIEZ",
        "ONCE",
        "DOCE",
        "TRECE",
        "CATORCE",
        "QUINCE",
        "DIECISEIS",
        "DIECISIETE",
        "DIECIOCHO",
        "DIECINUEVE",
    ]
    decenas = ["", "", "VEINTE", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA"]
    centenas = [
        "",
        "CIENTO",
        "DOSCIENTOS",
        "TRESCIENTOS",
        "CUATROCIENTOS",
        "QUINIENTOS",
        "SEISCIENTOS",
        "SETECIENTOS",
        "OCHOCIENTOS",
        "NOVECIENTOS",
    ]

    def under_100(x: int) -> str:
        if x < 20:
            return unidades[x]
        if x < 30:
            if x == 20:
                return "VEINTE"
            return ("VEINTI" + unidades[x - 20].lower()).upper()
        d, u = divmod(x, 10)
        if u == 0:
            return decenas[d]
        return f"{decenas[d]} Y {unidades[u]}"

    def under_1000(x: int) -> str:
        if x == 0:
            return ""
        if x == 100:
            return "CIEN"
        c, r = divmod(x, 100)
        if c == 0:
            return under_100(r)
        if r == 0:
            return centenas[c]
        return f"{centenas[c]} {under_100(r)}"

    def words(x: int) -> str:
        if x == 0:
            return "CERO"
        parts: list[str] = []
        millones, rem = divmod(x, 1_000_000)
        miles, resto = divmod(rem, 1000)
        if millones:
            if millones == 1:
                parts.append("UN MILLON")
            else:
                parts.append(f"{words(millones)} MILLONES")
        if miles:
            if miles == 1:
                parts.append("MIL")
            else:
                parts.append(f"{under_1000(miles)} MIL")
        if resto:
            parts.append(under_1000(resto))
        return " ".join(p for p in parts if p).upper()

    return f"SON PESOS {words(entero)} CON {cent:02d}/100"


def odoo_connect(cfg: dict) -> tuple[Any, int]:
    url = cfg["url"].rstrip("/")
    common = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/common", allow_none=True)
    uid = common.authenticate(cfg["db"], cfg["username"], cfg["password"], {})
    if not uid:
        raise SystemExit("Autenticación Odoo fallida (revisá ODOO_CONFIG_MASTER_DEV).")
    return xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/object", allow_none=True), int(uid)


def _convert_xlsm_to_xlsx(xlsm_path: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
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
        str(out_dir),
        str(xlsm_path),
    ]
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        raise SystemExit(f"Fallo conversión XLSM→XLSX.\nSTDOUT:\n{r.stdout}\nSTDERR:\n{r.stderr}\n")
    out = out_dir / (xlsm_path.stem + ".xlsx")
    if not out.exists():
        raise SystemExit(f"No se encontró el XLSX convertido: {out}")
    return out


def _xlsx_set_cells_local_sheet(template_xlsx: Path, *, out_xlsx: Path, cells: dict[str, tuple[str, str]]) -> Path:
    """
    Escribe celdas sobre la hoja `LOCAL` (sheet2.xml) porque ahí está el Print_Area.
    cells: { 'B15': ('s'|'n', value_as_str) }
    """
    out_xlsx.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(template_xlsx, "r") as z:
        names = z.namelist()

        shared: list[str] = []
        if "xl/sharedStrings.xml" in names:
            sst = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in sst.findall("m:si", NS):
                shared.append("".join((t.text or "") for t in si.findall(".//m:t", NS)))
        shared_index = {s: i for i, s in enumerate(shared)}

        def ensure_shared(s: str) -> int:
            if s in shared_index:
                return shared_index[s]
            shared_index[s] = len(shared)
            shared.append(s)
            return shared_index[s]

        def build_shared_xml() -> bytes:
            root = ET.Element(
                "{%s}sst" % NS["m"],
                {"count": str(len(shared)), "uniqueCount": str(len(shared))},
            )
            for s in shared:
                si = ET.SubElement(root, "{%s}si" % NS["m"])
                t = ET.SubElement(si, "{%s}t" % NS["m"])
                t.text = s
            return ET.tostring(root, encoding="utf-8", xml_declaration=True)

        sheet_path = "xl/worksheets/sheet2.xml"
        sh = ET.fromstring(z.read(sheet_path))
        sheetData = sh.find(".//m:sheetData", NS)
        if sheetData is None:
            raise SystemExit("Formato XLSX inesperado: falta sheetData.")

        rows_by_r: dict[int, ET.Element] = {}
        for row in sheetData.findall("m:row", NS):
            rr = int(row.attrib.get("r", "0"))
            rows_by_r[rr] = row

        def set_cell(cell_ref: str, kind: str, value: str) -> None:
            m = re.match(r"^([A-Z]+)(\d+)$", cell_ref.upper().strip())
            if not m:
                raise ValueError(f"Celda inválida: {cell_ref}")
            rnum = int(m.group(2))

            row = rows_by_r.get(rnum)
            if row is None:
                row = ET.SubElement(sheetData, "{%s}row" % NS["m"], {"r": str(rnum)})
                rows_by_r[rnum] = row

            target = None
            for c in row.findall("m:c", NS):
                if c.attrib.get("r") == cell_ref:
                    target = c
                    break
            if target is None:
                target = ET.SubElement(row, "{%s}c" % NS["m"], {"r": cell_ref})

            for ch in list(target):
                target.remove(ch)

            if kind == "s":
                target.attrib["t"] = "s"
                idx = ensure_shared(value)
                v = ET.SubElement(target, "{%s}v" % NS["m"])
                v.text = str(idx)
            else:
                target.attrib.pop("t", None)
                v = ET.SubElement(target, "{%s}v" % NS["m"])
                v.text = str(value)

        for ref, (kind, val) in cells.items():
            set_cell(ref, kind, val)

        # ordenar rows
        ordered_rows = sorted(rows_by_r.items(), key=lambda kv: kv[0])
        for row in list(sheetData.findall("m:row", NS)):
            sheetData.remove(row)
        for _, row in ordered_rows:
            sheetData.append(row)

        with zipfile.ZipFile(out_xlsx, "w") as outz:
            for n in names:
                if n == sheet_path:
                    outz.writestr(n, ET.tostring(sh, encoding="utf-8", xml_declaration=True))
                elif n == "xl/sharedStrings.xml":
                    outz.writestr(n, build_shared_xml())
                else:
                    outz.writestr(n, z.read(n))

    return out_xlsx


def _strip_local_drawings_and_comments(xlsx_path: Path) -> Path:
    """
    Quita objects de dibujo (textboxes/“notas guía”) y comentarios asociados a la hoja LOCAL (sheet2),
    para que no salgan impresos en el PDF.
    """
    sheet_path = "xl/worksheets/sheet2.xml"
    rels_path = "xl/worksheets/_rels/sheet2.xml.rels"
    rel_ns = {"pr": "http://schemas.openxmlformats.org/package/2006/relationships"}

    with zipfile.ZipFile(xlsx_path, "r") as z:
        names = set(z.namelist())
        if sheet_path not in names:
            return xlsx_path

        # 1) Editar sheet2.xml: remover nodos <drawing/> / legacyDrawing / comments.
        sh = ET.fromstring(z.read(sheet_path))
        # tags usuales:
        # - {..}drawing
        # - {..}legacyDrawing
        # - {..}legacyDrawingHF
        removed_any = False
        for tag_suffix in ("drawing", "legacyDrawing", "legacyDrawingHF", "comments"):
            for el in list(sh.findall(f".//m:{tag_suffix}", NS)):
                parent = sh
                # buscar parent manual (ET no lo provee); removemos en primer nivel si está directo
                # En sheet XML, estos nodos suelen ser hijos directos de <worksheet>.
                try:
                    sh.remove(el)
                    removed_any = True
                except Exception:
                    pass

        # 2) Editar rels para cortar vínculos y saber qué archivos borrar.
        drop_targets: set[str] = set()
        new_rels_bytes: bytes | None = None
        if rels_path in names:
            rels = ET.fromstring(z.read(rels_path))
            kept = []
            for rel in rels.findall("pr:Relationship", rel_ns):
                rtype = (rel.attrib.get("Type") or "").lower()
                tgt = rel.attrib.get("Target") or ""
                if "drawing" in rtype or "comments" in rtype or "vmlDrawing".lower() in rtype:
                    # guardar target relativo (ej. ../drawings/drawing1.xml)
                    drop_targets.add(tgt)
                else:
                    kept.append(rel)
            # reconstruir
            rels_new = ET.Element(rels.tag, rels.attrib)
            for rel in kept:
                rels_new.append(rel)
            new_rels_bytes = ET.tostring(rels_new, encoding="utf-8", xml_declaration=True)

        # normalizar targets a paths dentro del zip
        drop_paths: set[str] = set()
        for tgt in drop_targets:
            t = tgt.lstrip("/")
            if t.startswith("../"):
                t = t[3:]
            if not t.startswith("xl/"):
                t = "xl/" + t
            drop_paths.add(t)
            # algunos drawings traen rels propios (de imagenes), no hace falta borrar imágenes si no se usan;
            # pero borrar el drawing.xml alcanza para que no se renderice.

        # 3) Re-escribir ZIP filtrando entradas.
        tmp_bytes = {}
        tmp_bytes[sheet_path] = ET.tostring(sh, encoding="utf-8", xml_declaration=True) if removed_any else None
        if new_rels_bytes is not None:
            tmp_bytes[rels_path] = new_rels_bytes

        with tempfile.NamedTemporaryFile(prefix="strip_", suffix=".xlsx", delete=False) as tf:
            tmp_out = Path(tf.name)

        with zipfile.ZipFile(tmp_out, "w") as outz:
            for n in z.namelist():
                if n in drop_paths:
                    continue
                if n in tmp_bytes and tmp_bytes[n] is not None:
                    outz.writestr(n, tmp_bytes[n])
                    continue
                outz.writestr(n, z.read(n))

    tmp_out.replace(xlsx_path)
    return xlsx_path


def _add_signature_image_to_local_sheet(
    xlsx_path: Path,
    *,
    signature_png: Path,
    # Subimos un poco y achicamos para evitar 2da página.
    anchor_cell: str = "E42",
    width_emu: int = 1_650_000,
    height_emu: int | None = None,
) -> Path:
    """
    Inserta una imagen PNG (firma) en la hoja LOCAL usando OOXML drawings.
    Se hace después de strippear dibujos guía para no reintroducirlos.
    """
    if not signature_png.exists():
        return xlsx_path
    if height_emu is None:
        try:
            w, h = _png_wh(signature_png)
            # Mantener proporción (evita "deforme")
            height_emu = max(1, int(width_emu * h / max(1, w)))
        except Exception:
            height_emu = 600_000

    # convertir A1 ref a (col,row) 0-based
    m = re.match(r"^([A-Z]+)(\d+)$", anchor_cell.upper().strip())
    if not m:
        raise ValueError(f"Celda inválida para ancla: {anchor_cell}")
    col_letters, row_num_s = m.group(1), m.group(2)
    col = 0
    for ch in col_letters:
        col = col * 26 + (ord(ch) - ord("A") + 1)
    col -= 1
    row = int(row_num_s) - 1

    sheet_path = "xl/worksheets/sheet2.xml"
    rels_path = "xl/worksheets/_rels/sheet2.xml.rels"
    pkg_rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    rel_ns = {"pr": pkg_rel_ns}
    ws_ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    drawing_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:oneCellAnchor>
    <xdr:from>
      <xdr:col>{col}</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>{row}</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:ext cx="{cx}" cy="{cy}"/>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="1" name="firma.png"/>
        <xdr:cNvPicPr/>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip xmlns:r="{rns}" r:embed="rId1"/>
        <a:stretch><a:fillRect/></a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
</xdr:wsDr>
""".format(col=col, row=row, cx=width_emu, cy=height_emu, rns=ws_ns_r)

    drawing_rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    Target="../media/firma.png"/>
</Relationships>
"""

    with zipfile.ZipFile(xlsx_path, "r") as z:
        names = set(z.namelist())
        if sheet_path not in names:
            return xlsx_path

        # determinar próximo drawingN.xml disponible
        drawing_nums = []
        for n in names:
            m = re.match(r"xl/drawings/drawing(\d+)\.xml$", n)
            if m:
                drawing_nums.append(int(m.group(1)))
        next_n = (max(drawing_nums) + 1) if drawing_nums else 1
        drawing_path = f"xl/drawings/drawing{next_n}.xml"
        drawing_rels_path = f"xl/drawings/_rels/drawing{next_n}.xml.rels"

        # cargar y modificar sheet2.xml para apuntar al drawing
        ws = ET.fromstring(z.read(sheet_path))
        # agregar <drawing r:id="rIdSig"/> como hijo directo si no existe
        has_drawing = any(child.tag.endswith("drawing") for child in list(ws))
        if not has_drawing:
            dr = ET.Element("{%s}drawing" % NS["m"])
            dr.set("{%s}id" % ws_ns_r, "rIdSig")
            ws.append(dr)

        # actualizar/crear rels de sheet2 para agregar relationship a drawing
        rels_bytes = None
        if rels_path in names:
            rels = ET.fromstring(z.read(rels_path))
        else:
            rels = ET.Element("{%s}Relationships" % pkg_rel_ns)
        # evitar duplicado
        exists = False
        for rel in rels.findall("pr:Relationship", rel_ns):
            if rel.attrib.get("Id") == "rIdSig":
                exists = True
                rel.attrib["Target"] = f"../drawings/drawing{next_n}.xml"
        if not exists:
            rel = ET.SubElement(rels, "{%s}Relationship" % pkg_rel_ns)
            rel.attrib["Id"] = "rIdSig"
            rel.attrib["Type"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
            rel.attrib["Target"] = f"../drawings/drawing{next_n}.xml"
        rels_bytes = ET.tostring(rels, encoding="utf-8", xml_declaration=True)

        # asegurar content-types para drawings
        ct_path = "[Content_Types].xml"
        ct_bytes: bytes | None = None
        if ct_path in names:
            ct = ET.fromstring(z.read(ct_path))
            ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
            # root suele venir ya namespaced; buscamos Override existente
            drawing_part = f"/xl/drawings/drawing{next_n}.xml"
            has_override = False
            for ov in ct.findall(".//{*}Override"):
                if ov.attrib.get("PartName") == drawing_part:
                    has_override = True
                    break
            if not has_override:
                ET.SubElement(
                    ct,
                    "{%s}Override" % ct_ns,
                    {
                        "PartName": drawing_part,
                        "ContentType": "application/vnd.openxmlformats-officedocument.drawing+xml",
                    },
                )
            ct_bytes = ET.tostring(ct, encoding="utf-8", xml_declaration=True)

        # escribir zip nuevo
        with tempfile.NamedTemporaryFile(prefix="sig_", suffix=".xlsx", delete=False) as tf:
            tmp_out = Path(tf.name)

        firma_bytes = signature_png.read_bytes()
        with zipfile.ZipFile(tmp_out, "w") as outz:
            for n in z.namelist():
                if n == sheet_path:
                    outz.writestr(n, ET.tostring(ws, encoding="utf-8", xml_declaration=True))
                elif n == rels_path:
                    outz.writestr(n, rels_bytes)
                elif n == ct_path and ct_bytes is not None:
                    outz.writestr(n, ct_bytes)
                else:
                    outz.writestr(n, z.read(n))
            # agregar drawing + rels + imagen
            outz.writestr(drawing_path, drawing_xml)
            outz.writestr(drawing_rels_path, drawing_rels_xml)
            outz.writestr("xl/media/firma.png", firma_bytes)

    tmp_out.replace(xlsx_path)
    return xlsx_path


def _xlsx_to_pdf(xlsx_path: Path, out_pdf: Path) -> Path:
    out_pdf.parent.mkdir(parents=True, exist_ok=True)
    tmp_dir = out_pdf.parent
    cmd = [
        "libreoffice",
        "--headless",
        "--nologo",
        "--nolockcheck",
        "--nodefault",
        "--norestore",
        "--convert-to",
        "pdf",
        "--outdir",
        str(tmp_dir),
        str(xlsx_path),
    ]
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        raise SystemExit(f"Fallo conversión XLSX→PDF.\nSTDOUT:\n{r.stdout}\nSTDERR:\n{r.stderr}\n")
    produced = tmp_dir / (xlsx_path.stem + ".pdf")
    if not produced.exists():
        raise SystemExit(f"LibreOffice no generó PDF esperado: {produced}")
    if produced.resolve() != out_pdf.resolve():
        produced.replace(out_pdf)
    return out_pdf


def _parse_invoice_from_move_name(move_name: str) -> str:
    # Ej: 'PGAL1/... (FC 011-00573957)' -> '0011-00573957'
    m = re.search(r"\((?:FC\s+)?(\d{1,4})-(\d{1,10})\)", move_name or "", flags=re.IGNORECASE)
    if not m:
        return ""
    pv = m.group(1).zfill(4)[-4:]
    nro = m.group(2).zfill(8)[-8:]
    return f"{pv}-{nro}"


_PVNRO_ANY_RE = re.compile(r"(\d{1,4})-(\d{1,10})")


def _best_comprobante_from_memo_or_bills(memo: str, bills: list[dict]) -> str:
    """
    Prioridad pedida por estudio:
    - Usar `account.payment.memo` si contiene un comprobante (ej. "FC 011-00573957").
    - Si el memo es tipo BATCH/… o no trae PV-NRO, usar `ref` de facturas reconciliadas.
    """
    mm = (memo or "").strip()
    m = _PVNRO_ANY_RE.search(mm)
    if m:
        return mm
    # fallback: buscar ref con PV-NRO
    vals: list[str] = []
    for b in bills:
        ref = str(b.get("ref") or "").strip()
        if not ref:
            continue
        if _PVNRO_ANY_RE.search(ref):
            vals.append(ref)
    # si hay una sola, perfecto; si hay varias, concatenar (la plantilla tiene 1 campo)
    if not vals:
        return mm
    if len(vals) == 1:
        return vals[0]
    # Pedido: listar todos separados por coma (en una sola línea).
    return ", ".join(vals)


def _png_wh(path: Path) -> tuple[int, int]:
    """
    Devuelve (width,height) leyendo el chunk IHDR de un PNG (sin dependencias externas).
    """
    b = path.read_bytes()
    if len(b) < 24 or b[:8] != b"\x89PNG\r\n\x1a\n":
        raise ValueError(f"No parece PNG: {path}")
    if b[12:16] != b"IHDR":
        raise ValueError(f"PNG inválido (sin IHDR): {path}")
    w = int.from_bytes(b[16:20], "big")
    h = int.from_bytes(b[20:24], "big")
    return w, h


def _safe_filename_component(s: str, *, max_len: int = 80) -> str:
    s = (s or "").strip()
    if not s:
        return "SIN_NOMBRE"
    # quitar tildes/diacríticos
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.upper()
    s = re.sub(r"[^A-Z0-9]+", "_", s).strip("_")
    if not s:
        s = "SIN_NOMBRE"
    return s[:max_len]


def _last5_from_cert(cert_nro: str) -> str:
    d = _digits_only(cert_nro)
    return d[-5:] if len(d) >= 5 else d.rjust(5, "0")


def main(argv: Iterable[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--desde", required=True, help="YYYY-MM-DD")
    ap.add_argument("--hasta", required=True, help="YYYY-MM-DD")
    ap.add_argument(
        "--plantilla",
        type=Path,
        default=Path(__file__).resolve().parent / "PLANTILLA DE RETENCIONES-.xlsm",
    )
    ap.add_argument("--out-dir", type=Path, default=Path(__file__).resolve().parent / "out")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument(
        "--firma",
        type=Path,
        default=Path(__file__).resolve().parent / "firma.png",
        help="PNG de firma a insertar (opcional).",
    )
    ap.add_argument(
        "--keep-xlsx",
        action="store_true",
        help="Deja el XLSX intermedio (debug) en out-dir.",
    )
    args = ap.parse_args(list(argv) if argv is not None else None)

    # localizar config_nakel desde ARCA-RETENCIONES
    arca_root = Path(__file__).resolve()
    while arca_root != arca_root.parent:
        if (arca_root / "ARCA-RETENCIONES" / "SICORE" / "run_quincena.py").is_file():
            break
        arca_root = arca_root.parent
    sys.path.insert(0, str(arca_root / "ARCA-RETENCIONES"))
    from nakel_import_paths import prepend_config_nakel_sys_path  # type: ignore

    prepend_config_nakel_sys_path(arca_root / "ARCA-RETENCIONES")
    from config_nakel import ODOO_CONFIG_MASTER_DEV  # type: ignore

    models, uid = odoo_connect(ODOO_CONFIG_MASTER_DEV)
    db, pwd = ODOO_CONFIG_MASTER_DEV["db"], ODOO_CONFIG_MASTER_DEV["password"]

    dom = [
        ("date", ">=", args.desde),
        ("date", "<=", args.hasta),
        ("name", "ilike", "CERT-"),
        ("tax_line_id", "!=", False),
        ("parent_state", "=", "posted"),
    ]
    line_ids: list[int] = models.execute_kw(
        db, uid, pwd, "account.move.line", "search", [dom], {"order": "date asc, id asc"}
    )
    if args.limit and args.limit > 0:
        line_ids = line_ids[: args.limit]
    if not line_ids:
        raise SystemExit("No se encontraron líneas CERT-* con tax_line_id en el rango.")

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
                "name",
                "date",
                "partner_id",
                "move_id",
                "payment_id",
                "tax_line_id",
                "tax_base_amount",
                "balance",
                "credit",
            ]
        },
    )

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
            {
                "fields": [
                    "id",
                    "name",
                    "vat",
                    "l10n_ar_vat",
                    "street",
                    "street2",
                    "city",
                    "zip",
                    "l10n_ar_gross_income_number",
                ]
            },
        )
        partners = {int(r["id"]): r for r in pr}

    tax_ids = sorted({l["tax_line_id"][0] for l in lines if l.get("tax_line_id")})
    taxes: dict[int, dict] = {}
    if tax_ids:
        tr = models.execute_kw(db, uid, pwd, "account.tax", "read", [tax_ids], {"fields": ["id", "name", "amount"]})
        taxes = {int(r["id"]): r for r in tr}

    move_ids = sorted({l["move_id"][0] for l in lines if l.get("move_id")})
    moves: dict[int, dict] = {}
    if move_ids:
        mr = models.execute_kw(db, uid, pwd, "account.move", "read", [move_ids], {"fields": ["id", "name", "ref"]})
        moves = {int(r["id"]): r for r in mr}

    pay_ids = sorted({l["payment_id"][0] for l in lines if l.get("payment_id")})
    pays: dict[int, dict] = {}
    if pay_ids:
        pr = models.execute_kw(
            db,
            uid,
            pwd,
            "account.payment",
            "read",
            [pay_ids],
            {"fields": ["id", "memo", "reconciled_bill_ids"]},
        )
        pays = {int(r["id"]): r for r in pr}

    bill_ids = sorted({bid for p in pays.values() for bid in (p.get("reconciled_bill_ids") or [])})
    bills_by_id: dict[int, dict] = {}
    if bill_ids:
        br = models.execute_kw(db, uid, pwd, "account.move", "read", [bill_ids], {"fields": ["id", "ref", "name", "move_type"]})
        bills_by_id = {int(b["id"]): b for b in br}

    out_dir: Path = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="certpdf_") as td:
        tmp = Path(td)
        template_xlsx = _convert_xlsm_to_xlsx(Path(args.plantilla), tmp)

        for l in lines:
            cert_nro = str(l.get("name") or "").strip()
            emision = datetime.strptime(str(l.get("date")), "%Y-%m-%d").date()
            partner = partners.get(int(l["partner_id"][0])) if l.get("partner_id") else {}
            tax = taxes.get(int(l["tax_line_id"][0])) if l.get("tax_line_id") else {}
            move = moves.get(int(l["move_id"][0])) if l.get("move_id") else {}
            pay = pays.get(int(l["payment_id"][0])) if l.get("payment_id") else {}

            base = _abs_money(l.get("tax_base_amount"))
            importe = _abs_money(l.get("credit") or l.get("balance") or 0)
            ali_pct = _d2(tax.get("amount")).quantize(DEC2, rounding=ROUND_HALF_UP)
            ali_frac = (ali_pct / Decimal("100")).quantize(Decimal("0.0000001"), rounding=ROUND_HALF_UP)

            prov_name = str(partner.get("name") or "")
            prov_cuit = _cuit11_from_partner(partner)
            prov_iibb = str(partner.get("l10n_ar_gross_income_number") or "")
            prov_dom = " ".join(x for x in [partner.get("street"), partner.get("street2")] if x)
            prov_city = str(partner.get("city") or "")
            prov_zip = str(partner.get("zip") or "")

            memo = str(pay.get("memo") or "").strip()
            bill_list = [bills_by_id.get(int(bid), {}) for bid in (pay.get("reconciled_bill_ids") or [])]
            bill_list = [b for b in bill_list if b]
            nro_comp = _best_comprobante_from_memo_or_bills(memo, bill_list) or _parse_invoice_from_move_name(str(move.get("name") or "")) or str(move.get("ref") or "")
            letras = _num_to_words_es_pesos(importe)

            cells: dict[str, tuple[str, str]] = {
                "F6": ("s", cert_nro),
                "F7": ("n", str(_excel_date_serial(emision))),
                "B15": ("s", prov_name),
                "B16": ("s", prov_cuit),
                "B17": ("s", prov_iibb),
                "B18": ("s", prov_dom),
                "B19": ("s", prov_city),
                "B20": ("s", prov_zip),
                "B24": ("s", nro_comp),
                "B25": ("n", str(_excel_date_serial(emision))),
                "B26": ("n", str(base)),
                "B27": ("s", str(tax.get("name") or "")),
                "D31": ("n", str(base)),
                "E31": ("n", str(ali_frac)),
                "F31": ("n", str(importe)),
                "F32": ("n", str(importe)),
                "C34": ("n", str(importe)),
                "A36": ("s", letras),
            }

            out_xlsx = tmp / f"{cert_nro}.xlsx"
            _xlsx_set_cells_local_sheet(template_xlsx, out_xlsx=out_xlsx, cells=cells)
            _strip_local_drawings_and_comments(out_xlsx)
            _add_signature_image_to_local_sheet(
                out_xlsx,
                signature_png=args.firma,
            )
            out_pdf = out_dir / f"CERT-{_safe_filename_component(prov_name)}-{_last5_from_cert(cert_nro)}.pdf"
            _xlsx_to_pdf(out_xlsx, out_pdf)
            if args.keep_xlsx:
                kept = out_dir / f"{out_pdf.stem}.xlsx"
                shutil.copyfile(out_xlsx, kept)

    print(f"OK: generados {len(lines)} PDF en {out_dir}")
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())

