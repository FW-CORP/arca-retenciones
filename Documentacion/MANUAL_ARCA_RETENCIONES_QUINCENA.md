---
title: ARCA Retenciones — manual de uso (quincena / rango)
updated: 2026-04-20
---

## Para qué sirve

Este proyecto genera archivos para **ARCA / certificaciones** a partir de retenciones registradas en **Odoo** (solo lectura por XML-RPC). Lo habitual es un **rango de fechas** (por ejemplo quincena: `2026-04-01` a `2026-04-15`).

## Qué tenés que usar el 99% del tiempo

| Objetivo | Comando (desde la carpeta `arca-retenciones/`) | Archivo generado |
|----------|-----------------------------------------------|------------------|
| **SICORE — Ganancias** (TXT importación) | `python3 SICORE/run_quincena.py --desde YYYY-MM-DD --hasta YYYY-MM-DD` | `SICORE/out/SICORE_V9_RET_GAN_YYYY-MM-DD_a_YYYY-MM-DD.TXT` |
| **SIRCAR — IIBB** (TXT 163 posiciones) | `python3 SIRCAR/run_quincena.py --desde YYYY-MM-DD --hasta YYYY-MM-DD --cuit-agente 30XXXXXXXXX` | `SIRCAR/out/SIRCAR_163_YYYY-MM-DD_a_YYYY-MM-DD.TXT` |

### Ejemplos (quincenas)

```bash
cd /media/klap/raid5/cursor_files/nakel/arca-retenciones

# 1 al 15 de abril
python3 SICORE/run_quincena.py --desde 2026-04-01 --hasta 2026-04-15

# 16 al 30 de abril
python3 SICORE/run_quincena.py --desde 2026-04-16 --hasta 2026-04-30
```

SIRCAR (mismo criterio de fechas; el CUIT es el del **agente de retención**):

```bash
python3 SIRCAR/run_quincena.py --desde 2026-04-01 --hasta 2026-04-15 --cuit-agente 30500000000
```

## Contra qué base de Odoo corre

Por defecto los scripts usan `config_nakel.ODOO_CONFIG_MASTER_DEV` (normalmente **producción**: `nakel.net.ar` / `master_dev`).

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

Las salidas por defecto de esos scripts van a **`SICORE/out/`** o **`SIRCAR/out/`** (no a `tools/out/`).

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
