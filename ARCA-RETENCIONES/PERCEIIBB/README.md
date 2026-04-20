---
title: PERCEIIBB — Percepciones IIBB ARCA (TXT 163)
updated: 2026-04-20
---

## Ubicación

Esta carpeta vive **dentro de** `ARCA-RETENCIONES/PERCEIIBB/` (mismo monorepo que retenciones SICORE/SIRCAR).

## Objetivo

Export **TXT ancho fijo 163** para **percepciones IIBB** (SIRCAR / ARCA), según la tabla del estudio:

- **Campo 1** = `2` (percepción; en retenciones va `1`).
- **Fecha** = **fecha de factura de venta** (`invoice_date`), no fecha de pago.
- **Campo 8** = `F` / `D` / `C` (factura, nota de débito, nota de crédito). En NC el importe percibido va en **positivo**; el aplicativo resta al leer `C`.
- Origen: facturas de cliente publicadas en el rango + líneas `account.move.line` con `tax_line_id` en impuestos cuyo `l10n_ar_tax_type` contiene **`perception`** (y fallback por nombre `PERC` + IIBB).

## Configuración

Los scripts localizan la raíz `ARCA-RETENCIONES/` (marcadores `nakel_import_paths.py` + `SICORE/run_quincena.py`) y cargan **`config_nakel`** igual que el resto del proyecto.

## Uso

Desde la carpeta **`ARCA-RETENCIONES/`**:

```bash
python3 PERCEIIBB/run_quincena.py --desde 2026-04-01 --hasta 2026-04-15 --cuit-agente 30XXXXXXXXX
```

Salida: `PERCEIIBB/out/PERCEIIBB_ARCA_YYYY-MM-DD_a_YYYY-MM-DD.TXT`

Opciones extra (`--jurisdiccion-sujeto`, `--cuota`, etc.) se reenvían al generador.

## Referencia de layout

Tabla posicional provista por el estudio (percepciones 163); misma longitud que retenciones SIRCAR con la semántica indicada arriba.
