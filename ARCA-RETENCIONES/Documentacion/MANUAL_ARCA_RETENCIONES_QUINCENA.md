---
title: ARCA Retenciones — manual de uso (quincena / rango)
updated: 2026-04-21
---

## Para qué sirve

Este proyecto genera archivos para **ARCA / certificaciones** a partir de retenciones registradas en **Odoo** (solo lectura por XML-RPC). Lo habitual es un **rango de fechas** (por ejemplo quincena: `2026-04-01` a `2026-04-15`).

## Qué tenés que usar el 99% del tiempo

| Objetivo | Comando (desde la carpeta del clon `nakel_scripts/ARCA-RETENCIONES/`) | Archivo generado |
|----------|-----------------------------------------------|------------------|
| **SICORE — Ganancias** (TXT importación) | `python3 SICORE/run_quincena.py --desde YYYY-MM-DD --hasta YYYY-MM-DD` | `SICORE/out/SICORE_V9_RET_GAN_YYYY-MM-DD_a_YYYY-MM-DD.TXT` |
| **SIRCAR — IIBB** (TXT 163 posiciones) | `python3 SIRCAR/run_quincena.py --desde YYYY-MM-DD --hasta YYYY-MM-DD --cuit-agente 30XXXXXXXXX` | `SIRCAR/out/SIRCAR_163_YYYY-MM-DD_a_YYYY-MM-DD.TXT` |
| **ARCA — Percepciones IIBB** (TXT 163; campo 1 = `2`) | Ver ejemplo abajo (cwd = **raíz** del clon `arca-retenciones/`) | `PERCEIIBB/out/PERCEIIBB_ARCA_YYYY-MM-DD_a_YYYY-MM-DD.TXT` |

**Nota cwd**: SICORE y SIRCAR suelen ejecutarse desde `ARCA-RETENCIONES/`. **PERCEIIBB** está en la **raíz del clon** (al lado de `ARCA-RETENCIONES/`): ahí conviven las carpetas `SICORE/`, `SIRCAR/` y `PERCEIIBB/`.

### Ejemplos (quincenas)

```bash
cd /media/klap/raid5/cursor_files/nakel/nakel_scripts/ARCA-RETENCIONES

# 1 al 15 de abril
python3 SICORE/run_quincena.py --desde 2026-04-01 --hasta 2026-04-15

# 16 al 30 de abril
python3 SICORE/run_quincena.py --desde 2026-04-16 --hasta 2026-04-30
```

SIRCAR (mismo criterio de fechas; el CUIT es el del **agente de retención**):

```bash
python3 SIRCAR/run_quincena.py --desde 2026-04-01 --hasta 2026-04-15 --cuit-agente 30500000000
```

Percepciones IIBB (fecha según **factura de venta** en el rango; CUIT del **agente de percepción**):

```bash
cd …/arca-retenciones
python3 PERCEIIBB/run_quincena.py --desde 2026-04-01 --hasta 2026-04-15 --cuit-agente 30500000000
```

## Contra qué base de Odoo corre

Por defecto los scripts usan `config_nakel.ODOO_CONFIG_MASTER_DEV` (normalmente **producción**: `nakel.net.ar` / `master_dev`).

Si en otra PC falla el import de `config_nakel`, definí **`NAKEL_CONFIG_ROOT`** apuntando al directorio que contiene `config_nakel.py` (ver README de `ARCA-RETENCIONES`).

Para apuntar a **desarrollo** (`dev.nakel.net.ar` / `master_test`), exportá antes:

```bash
export NAKEL_TARGET=master_test
python3 SICORE/run_quincena.py --desde 2026-04-01 --hasta 2026-04-15
```

Las credenciales de `master_test` suelen venir de `nakel/.env` (`ODOO_MASTER_DEV_*`), que `config_nakel.py` carga automáticamente.

## Validación automática (SICORE)

`SICORE/run_quincena.py` llama después a `--validar-posiciones` sobre el TXT generado. Si falla, revisá el mensaje (líneas cortas o campos desalineados).

Para omitir esa validación:

```bash
python3 SICORE/run_quincena.py --desde 2026-04-01 --hasta 2026-04-15 --skip-validacion
```

## Herramientas extra (no son el flujo principal)

Están en subcarpetas `tools/` para no mezclarlas con el generador principal:

| Carpeta | Scripts | Uso típico |
|---------|---------|------------|
| `SICORE/tools/` | `generar_op_odoo_master_dev.py`, `generar_ret_gan_mayor_odoo_master_dev.py`, `generar_rgan_cpa_master_dev.py` | Planillas / mayores / layout alternativo Ganancias |
| `SIRCAR/tools/` | `generar_sircar_mayor_odoo_master_dev.py`, `generar_ret_iibb_mayor_odoo_master_dev.py`, `generar_ret_dgr_*.py` | CSV mayor, mayor IIBB, exports DGR |
| `CERTIFICADOS-RETENCION-PDF/` | `generar_certificados_retencion_pdf_master_dev.py` | Generación masiva de **certificados PDF** (plantilla Excel → PDF) |

Las salidas por defecto de esos scripts van a **`SICORE/out/`** o **`SIRCAR/out/`** (no a `tools/out/`).

### Certificados de retención (PDF)

Se generan desde la carpeta `CERTIFICADOS-RETENCION-PDF/` (en la **raíz del clon** `arca-retenciones/`).

Ejemplo (quincena):

```bash
python3 CERTIFICADOS-RETENCION-PDF/generar_certificados_retencion_pdf_master_dev.py \
  --desde 2026-04-01 --hasta 2026-04-15
```

Firma digital (opcional):

- Por defecto toma `CERTIFICADOS-RETENCION-PDF/firma.png` y la inserta en la hoja `LOCAL` antes de exportar el PDF.
- Si querés otra imagen, usá `--firma /ruta/a/otra_firma.png`.

Debug:

- `--keep-xlsx` deja el `.xlsx` intermedio al lado del PDF, útil si hay que ajustar posición/tamaño de la firma.
  La firma se inserta manteniendo proporción (sin deformarse) y se ubica para que el PDF salga en 1 sola página.

## Parámetros avanzados

Los generadores “de verdad” siguen siendo `SICORE/generar_sicore_v9_retenciones.py` y `SIRCAR/generar_sircar_163_master_dev.py`.  
`run_quincena.py` reenvía al script subyacente **cualquier flag que no reconozca** (por ejemplo `--codigo-condicion`, `--modo-comprobante`, etc.).

Ejemplo (forzar código de comprobante en SICORE):

```bash
python3 SICORE/run_quincena.py --desde 2026-04-01 --hasta 2026-04-15 --codigo-comprobante 06
```

(No dupliqués `--desde`, `--hasta`, `--out` ni `--codigo-operacion` si ya los usás en el wrapper; si necesitás control total, llamá directo al `generar_*.py`.)

## Documentación técnica

- Especificación SICORE y mapeo Odoo: `Documentacion/ARCA_RETENCIONES_LAYOUTS_Y_MAPEO_ODOO.md`
- SIRCAR (layout 163): `SIRCAR/README.md`
