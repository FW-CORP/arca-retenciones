---
title: ARCA Retenciones — layouts y mapeo a Odoo (solo lectura)
updated: 2026-04-23
---

## Objetivo

Documentar de forma **canónica** cómo se construyen los TXT de retenciones usados por el estudio/contador para ARCA/DGR/SIRCAR/Ganancias, y cómo obtener la información desde **Odoo `master_dev`** mediante **API XML-RPC en modo solo consulta**.

Este documento unifica:
- La especificación de **SICORE 9.0** (`arca-doc.md`).
- Los **layouts reales** observados en los TXT del directorio `ARCA-RETENCIONES/`.
- El **mapeo de datos** contra Odoo 18 Argentina (retenciones registradas en pagos).

---

## Fuente de verdad (ARCA / SICORE)

Ver `ARCA-RETENCIONES/Documentacion/arca-doc.md` (SICORE 9.0). Puntos críticos:

- **Encoding**: ASCII/ANSI.
- **Fin de línea**: **CRLF** (Windows) al final de cada registro.
- **Separador decimal**: **coma** (configurable en el aplicativo).
- **Alineación**:
  - alfanumérico: **izquierda**, relleno con espacios
  - numérico: **derecha**, relleno con ceros (o espacios según el diseño)
- **Fecha**: `dd/mm/aaaa`.
- **CUIT**: sin guiones (según SICORE; en nuestros layouts aparecen variantes).
- **Longitud fija**: *todos los registros deben medir lo mismo*.

### Registro estándar SICORE (retenciones)

`arca-doc.md` detalla el registro estándar de **132 posiciones** (sin exterior) con campos como:
- Código de comprobante (2)
- Fecha comprobante (10)
- Nro comprobante (16)
- Importe comprobante (16)
- Código impuesto (4) (**cambio v9**)
- Código régimen (4)
- Código operación (1)
- Base (14)
- Fecha retención (10)
- Código condición (2) (**nuevo v9**)
- Sujeto suspendido (1)
- Importe retención (14)
- etc.

**Nota**: Los TXT locales (`RET-DGR.TXT`, `RET-DGR-SIRCAR.TXT`, `RGAN_CPA.TXT`) no son “idénticos” al estándar de 132; parecen ser **layouts de importación/salida del estudio** (o pre-SICORE / parametrizados). Por eso, la estrategia es: **replicar los layouts reales** y validar por importación.

### ARCA — Percepciones IIBB (TXT 163)

Carpeta **`PERCEIIBB/` en la raíz del clon** (salida `PERCEIIBB/out/`): mismo **ancho 163** que SIRCAR retenciones, con **tipo registro `2`**, fecha según **factura de venta**, tipo comprobante **F/D/C** y líneas de impuesto de **percepción** en facturas de cliente (`account.move` + `account.move.line.tax_line_id`). Detalle: `PERCEIIBB/README.md` y `generar_perceiibb_arca_master_dev.py`.

---

## Inventario de archivos y su rol

En `ARCA-RETENCIONES/` se manejan 3 TXT (muestras del contador):

1) **`RET-DGR-SIRCAR.TXT`** (CSV)
- 1 línea = 1 retención.
- 11 columnas, coma como separador.
- Decimal con **punto** (en la muestra del contador).

2) **`RET-DGR.TXT`** (ancho fijo)
- 1 **header** + N **detalles**
- Detalle incluye: identificador, fecha, importes serializados (centavos), CUIT con guiones, y datos del proveedor (nombre/domicilio/localidad/CP/provincia).

3) **`RGAN_CPA.TXT`** (ancho fijo)
- 1 línea = 1 retención.
- 145 caracteres por línea (en la muestra).
- Montos con **coma decimal**, fechas repetidas, CUIT como `80` + CUIT(11) (13 dígitos).

---

## Dónde están las retenciones en Odoo (master_dev)

En Odoo 18 Argentina, las retenciones se encuentran asociadas al pago:

- **Pago**: `account.payment`
  - **retenciones**: `l10n_ar_withholding_ids` (one2many a `account.move.line`)

- **Línea contable de retención**: `account.move.line`
  - **base**: `tax_base_amount`
  - **importe retenido**: suele venir en `credit` (y/o `balance`)
  - **impuesto/régimen**: `tax_line_id` → `account.tax`

- **Impuesto**: `account.tax`
  - `l10n_ar_tax_type`:
    - `earnings` → **Ganancias** (RG 830 / etc.)
    - `iibb_*` → **IIBB / SIRCAR** (según configuración)
  - `amount` → alícuota (%)
  - `l10n_ar_code` → código AFIP (si aplica)

**Clave**: para exportar TXT, se recomienda tomar los datos desde `account.payment` + `l10n_ar_withholding_ids` + `tax_line_id`.

---

## Mapeo por archivo (conceptual)

### A) `RET-DGR-SIRCAR.TXT` (CSV)

Campos típicos:
- correlativo (5 dígitos)
- lote/sub-lote (constantes)
- nro “orden” / correlativo interno (en el estudio aparece `000000018150`)
- CUIT proveedor (11)
- fecha (pago/retención)
- base imponible
- alícuota
- importe retenido
- códigos extra (ej. `001`, `907`)

**Origen en Odoo**:
- fecha: `account.payment.date`
- nro de comprobante/OP (recomendado post-migración): `account.payment.name` **normalizado sin separadores** (ej. `PGAL1/26-27/0370` → `PGAL126270370`)
- CUIT: `res.partner.l10n_ar_vat`/`vat`
- base: `account.move.line.tax_base_amount`
- importe: `account.move.line.credit` (abs)
- alícuota: `account.tax.amount` (del `tax_line_id`)

### B) `RET-DGR.TXT` (ancho fijo, header + detalle)

Detalle observado (longitudes reales del layout del contador):
- bloque importes serializados en centavos (3 importes de 12 dígitos)
- bloque CUIT `000-000000-0DD-XXXXXXXX-X`
- bloques de 30/30/20/8/2 para datos del proveedor

**Origen en Odoo**:
- “importe total pago”: `account.payment.amount` (abs) → centavos
- base: `tax_base_amount` → centavos
- importe retenido: `credit/balance` → centavos
- nro_orden (si el estudio acepta numeración nueva): `account.payment.name` normalizado sin separadores (adaptado al ancho del layout)
- datos proveedor: `res.partner.name/street/city/zip/state_id`
- CUIT con guiones: desde `l10n_ar_vat`/`vat` formateado

### C) `RGAN_CPA.TXT` (ancho fijo 145)

Este layout se parece al caso SICORE/Ganancias (código impuesto 0217) pero con campos del estudio:

- **Terminador de línea**: usar **CRLF** (`\r\n`) como en las muestras (SIAP/SICORE puede rechazar el archivo si no coincide el “largo de registro” esperado).
- **`nro_orden` (pos. 17–28, 12 chars)**: se compone como **PV(4) + NRO(8)** tomado de la **factura proveedor** (`account.move.ref` tipo `0388-00386471`). Si no hay `PV-NRO` parseable, se usa fallback numérico.
- `codigo_8` observado: `02170781`
- `jurisd_3` observado: `010`
- `cuit13` observado: `80` + CUIT(11)

**Origen en Odoo**:
- mismas fuentes de retención, pero filtrando `account.tax.l10n_ar_tax_type == 'earnings'`.
- nro_orden/nro_comprobante (si el estudio lo acepta): `account.payment.name` normalizado sin `/` ni `-` (y recortado/padded al ancho del layout).

### D) SIRCAR (ancho fijo 163)

Layout canónico (tabla completa en la captura):

- `assets/image-df88190b-8204-479b-840d-109d33628030.png`

Puntos críticos:
- **Longitud fija**: 163 caracteres por línea + CRLF.
- **Jurisdicción (campo 2)**: obligatoria (ej. 901 CABA, 902 BsAs… según tabla del estudio).
- **Cuota (campo 5)**: 1 o 2 (quincena).
- **Alícuota (campo 12)**: 5 dígitos con 3 decimales (sin coma).
- **Razón social (campo 17)**: 40 caracteres, relleno con **espacios a derecha**.

---

## Implementación (solo lectura) en este repo/vault

**Uso operativo (quincena / rango)**: ver `Documentacion/MANUAL_ARCA_RETENCIONES_QUINCENA.md` (`SICORE/run_quincena.py`, `SIRCAR/run_quincena.py`).

Scripts ya creados (solo consulta a `master_dev`):

- `ARCA-RETENCIONES/SIRCAR/tools/generar_ret_dgr_master_dev.py`
  - Genera **CSV** tipo `RET-DGR-SIRCAR.TXT` en `ARCA-RETENCIONES/SIRCAR/out/RET-DGR.TXT` (CSV).

- `ARCA-RETENCIONES/SIRCAR/tools/generar_ret_dgr_ancho_fijo_master_dev.py`
  - Genera **ancho fijo** tipo `RET-DGR.TXT` en `ARCA-RETENCIONES/SIRCAR/out/RET-DGR.TXT`.

- `ARCA-RETENCIONES/SICORE/tools/generar_rgan_cpa_master_dev.py`
  - Genera **ancho fijo 145** tipo `RGAN_CPA.TXT` en `ARCA-RETENCIONES/SICORE/out/RGAN_CPA.TXT`.

- `ARCA-RETENCIONES/SIRCAR/tools/generar_sircar_mayor_odoo_master_dev.py`
  - Genera **extracto tipo mayor** para **IIBB / SIRCAR** en `ARCA-RETENCIONES/SIRCAR/out/` (CSV).

- `ARCA-RETENCIONES/SICORE/generar_sicore_v9_retenciones.py`
  - Genera **SICORE / importación retenciones Ganancias** en **`SICORE/out/SICORE_V9_RETENCIONES_GANANCIAS.TXT`**: registro de **159** posiciones + CRLF (cola 131–159; último carácter **un espacio** según grilla importación).
  - **N° comprobante (16)**: solo dígitos; **4** = punto de venta + **12** = número (sin guiones). Textos internos Odoo (FACOM…) no son válidos fiscalmente: se arman dígitos desde `2526-30010` o solo dígitos del dato.
  - **Impuesto + régimen**: `0217` (4) + régimen **3** dígitos (p. ej. `078`) desde `account.tax.l10n_ar_code` (últimos 3 dígitos numéricos). El estándar de 132 posiciones en `arca-doc.md` usa régimen en **4** posiciones; aquí se compacta a 3 para cerrar el registro con la cola del estudio.
  - **Comprobante 01 (factura)**: con `reconciled_bill_ids` usa **01** y datos de factura; sin factura, **06** y datos de pago (`--modo-comprobante` / `--codigo-comprobante`).
  - **Montos**: sin separadores; **últimos 2 = centavos**.
  - **Código de operación** (Tabla C `arca-doc.md`): por defecto **1** = Retención (`--codigo-operacion`). El valor **0** no figura en la tabla oficial.
  - **Tras importe retención**: bloque de **13 ceros** (LPAD; sin espacios en el medio) para alinear el **tipo de documento** en columnas **107–108** y el **CUIT** en **109–119** según importación SICORE. Los campos numéricos no usan espacios.
  - **Inspección**: `python3 SICORE/generar_sicore_v9_retenciones.py --desglosar SICORE/out/SICORE_V9_RETENCIONES_GANANCIAS.TXT` imprime cada campo con rango de posiciones (sin conectar a Odoo).
  - **Columnas clave**: `python3 SICORE/generar_sicore_v9_retenciones.py --validar-posiciones SICORE/out/SICORE_V9_RETENCIONES_GANANCIAS.TXT` comprueba en todas las líneas: col **67** inicio fecha retención, col **107** `80`, col **109** primer dígito CUIT (criterio Bloc de notas).

**Nomenclatura (criterio estudio / lo que suelen pedir)** — no mezclar con el “tipo de comprobante” de la factura:

- **`80`**: en SICORE es el **tipo de documento del sujeto retenido** (Tabla F: **CUIT**), posiciones **107–108** del registro (implementación actual). **No** es el código de comprobante de la operación (eso son pos. **1–2**: p. ej. `01` Factura, `06` Orden de pago).
- **“Campo 10” del diseño estándar** (lista de campos del importador): suele corresponder al **código de condición** (nuevo en v9), posiciones **77–78**, **2 caracteres**. Para **Inscripto** va **`01`**, no un solo carácter `1`. El generador ya usa por defecto `--codigo-condicion 01` y `--tipo-doc 80`.

### Referencia canónica (captura del estudio)

La captura más completa (layout de **159** posiciones) quedó guardada en:

- `assets/image-4bb11c16-87d6-4a2f-8bdf-d4756f00c189.png`

Usar esa imagen como **fuente de verdad** para posiciones 1–159 (en particular campos 11–15 y la cola 16–20).

### Mapa posicional SICORE 159 (implementación actual)

Orden de concatenación (posiciones **1-based** inclusive):

| Desde | Hasta | Largo | Campo |
|------:|------:|------:|--------|
| 1 | 2 | 2 | Código de comprobante |
| 3 | 12 | 10 | Fecha comprobante `dd/mm/aaaa` |
| 13 | 28 | 16 | Número de comprobante (solo dígitos) |
| 29 | 44 | 16 | Importe comprobante (entero en centavos) |
| 45 | 48 | 4 | Código de impuesto (`0217`) |
| 49 | 51 | 3 | Código de régimen |
| 52 | 52 | 1 | Código de operación |
| 53 | 66 | 14 | Base (centavos) |
| 67 | 76 | 10 | Fecha retención `dd/mm/aaaa` |
| 77 | 78 | 2 | Código condición (v9) |
| 79 | 79 | 1 | Retención sujeto suspendido |
| 80 | 93 | 14 | Importe retención (centavos) |
| 94 | 106 | 13 | Relleno numérico (`0`×13: % exclusión / ajuste / boletín según criterio importación; no espacios) |
| 107 | 108 | 2 | Tipo documento retenido (`80` = CUIT) |
| 109 | 130 | 22 | Documento retenido: CUIT 11 dígitos + 11 ceros (LPAD) |
| 131 | 140 | 10 | Fecha publicación certificado |
| 141 | 142 | 2 | Tipo régimen especial |
| 143 | 156 | 14 | Importe base exclusión |
| 157 | 158 | 2 | Tipo cuenta (`00`) |
| 159 | 159 | 1 | Relleno final (**un espacio**; no dos ceros) |

**Grilla importación (campos 15–20 de la planilla)**: certificado/exclusión 11 ceros + fecha `00/00/0000` + tipo régimen `00` + base exclusión 14 ceros + tipo cuenta `00` + **1 espacio** — equivalente a pos. 131–159 del mapa anterior (ajustar si la planilla numera columnas distinto).

**Valor retención (“campo 11” en planillas)**: el importe de la retención son **14** posiciones (80–93). La posición **79** es **sujeto suspendido** (1 dígito), no parte del importe; si la grilla agrupa 79–93 como un solo bloque visual, verá **15** caracteres sin que el importe tenga un dígito de más.

**Cadena “modelo” manual**: debe tener **159** caracteres útiles antes del CRLF. Si tiene menos (p. ej. 138) o incluye barras `/` entre bloques que no son fechas, los cortes fijos **desalinean** campos (por ejemplo el bloque `0217` puede leerse como `1707`). El número de comprobante fiscal no es un relleno arbitrario: para `2526-30010` el generador emite `2526000000030010` (PV cuatro dígitos + número doce dígitos), no patrones alternativos de 16 ceros mezclados con el número crudo.

### Validación rápida “llegar a 159” (regla visual)

Guía de Florencia (19/04/2026) para detectar líneas “cortas” o con corrimientos:

- **Hasta el CUIT (Campo 14)**: **119** caracteres.
- **Campo 15 (Exclusión / Certificado)**: **11** ceros (pos. **120–130**).
- **Campo 16 (Fecha Pub.)**: `00/00/0000` (pos. **131–140**).
- **Campo 17 al 19 (Rellenos)**: **18** ceros (pos. **141–158**).
- **Campo 20 (Final)**: **1 espacio** (pos. **159**).

- `ARCA-RETENCIONES/SICORE/tools/generar_op_odoo_master_dev.py`
  - Genera **`SICORE/out/OP_odoo.xlsx`** (misma grilla que `Documentacion/OP.xlsx`) desde `account.payment` + retenciones.
  - **Cheques de terceros vs efectivo**: Odoo puede marcar el diario “Cheques de Terceros” con **tipo cash**; para la planilla se usa **`account.journal.code`** (ej. `CHQS` → columna *Cheques Terceros*, `EFVO` → *Efectivo*) y el prefijo del pago (ej. `PCHQS/…`).

- `ARCA-RETENCIONES/SICORE/tools/generar_ret_gan_mayor_odoo_master_dev.py`
  - Genera **`SICORE/out/RET_GAN_odoo.xlsx`** (estructura tipo `Documentacion/RET GAN 16-03.xlsx`): **mayor de líneas** con `tax_line_id` en impuestos **`l10n_ar_tax_type = earnings`** (retención Ganancias), con **S.Ini.** = suma debe/haber antes del `--desde`.
  - Nota Odoo 18: esas líneas suelen tener `display_type = product` (no se filtra `display_type=False`).

- `ARCA-RETENCIONES/SIRCAR/tools/generar_ret_iibb_mayor_odoo_master_dev.py`
  - Genera **`SIRCAR/out/RET_IIBB_odoo.xlsx`** (estructura tipo `Documentacion/RETIIBB.xlsx`): **mayor de líneas** con `tax_line_id` en impuestos IIBB/SIRCAR (por `l10n_ar_tax_type ilike iibb` o `name ilike IIBB/SIRCAR`), con **S.Ini.** = suma debe/haber antes del `--desde`.

Todos:
- usan `config_nakel.ODOO_CONFIG_MASTER_DEV`
- conectan por XML-RPC
- hacen **solo `search_read`/`read`** (sin `write/create`)

---

## Checklist de validación con el contador

Antes de usar en producción:
- Confirmar **qué representa el correlativo** “orden de pago” del estudio (ej. `000000018150`).
  - En Odoo, `account.payment.name` es alfanumérico (`PGAL1/...`), así que ese correlativo sale de otro numerador del estudio o de otro sistema.
- Confirmar si los códigos constantes (`001`, `907`, `010`, `02170781`) son:
  - fijos por empresa,
  - fijos por régimen,
  - o parametrizables por jurisdicción.
- Importar un set de prueba en el aplicativo y revisar `errimpret.log` si rechaza.

