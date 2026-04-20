---
title: SICORE — Ganancias (solo lectura Odoo)
updated: 2026-04-20
---

## Objetivo

Generar el **TXT SICORE v9** (retenciones **Ganancias**) desde Odoo por **XML-RPC**, sin escribir en la base.

## Uso habitual (quincena o cualquier rango)

Desde la carpeta del clon `nakel_scripts/ARCA-RETENCIONES/`:

```bash
python3 SICORE/run_quincena.py --desde YYYY-MM-DD --hasta YYYY-MM-DD
```

- Salida: `SICORE/out/SICORE_V9_RET_GAN_YYYY-MM-DD_a_YYYY-MM-DD.TXT`
- Por defecto usa `--codigo-operacion 1` (Retención) y corre `--validar-posiciones` al final (omitir con `--skip-validacion`).

## Generador completo (todas las opciones)

```bash
python3 SICORE/generar_sicore_v9_retenciones.py --desde YYYY-MM-DD --hasta YYYY-MM-DD
```

## Herramientas extra

En `SICORE/tools/` (planillas / mayores / layout alternativo). Las salidas por defecto siguen yendo a **`SICORE/out/`**.

## Documentación

- Manual operativo: `Documentacion/MANUAL_ARCA_RETENCIONES_QUINCENA.md`
- Detalle técnico + mapeo Odoo: `Documentacion/ARCA_RETENCIONES_LAYOUTS_Y_MAPEO_ODOO.md`
