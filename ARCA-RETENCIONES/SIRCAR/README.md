---
title: SIRCAR — extracción desde Odoo (solo lectura)
updated: 2026-04-20
---

## Objetivo

Generar un extracto de **retenciones IIBB / SIRCAR** desde **Odoo `master_dev`** (XML-RPC) en **modo solo consulta**, en el formato requerido por el estudio/contador.

Este directorio agrupa todo lo relativo a SIRCAR para no mezclarlo con SICORE/Ganancias.

## Layout canónico (163 caracteres)

La especificación canónica (tabla de posiciones 1–163) está en la captura:

- `assets/image-df88190b-8204-479b-840d-109d33628030.png`

## Archivos

- `run_quincena.py`
  - Atajo: rango de fechas + CUIT agente → TXT con nombre estándar en `SIRCAR/out/`.

- `SIRCAR/tools/generar_sircar_mayor_odoo_master_dev.py`
  - Exporta un “mayor”/listado de líneas de retención IIBB/SIRCAR (baseado en `account.move.line` con `tax_line_id`).
  - Salida por defecto: `SIRCAR/out/SIRCAR_mayor.csv`

- `generar_sircar_163_master_dev.py`
  - Genera TXT **ancho fijo 163** según layout canónico (apto para upload SIRCAR).
  - Salida por defecto: `SIRCAR/out/SIRCAR_163.TXT`

## Cómo ejecutar

```bash
python3 SIRCAR/run_quincena.py --desde YYYY-MM-DD --hasta YYYY-MM-DD --cuit-agente 30XXXXXXXXX
```

Mayor / CSV (herramienta):

```bash
python3 SIRCAR/tools/generar_sircar_mayor_odoo_master_dev.py --desde YYYY-MM-DD --hasta YYYY-MM-DD
```

Para el layout 163 (generador completo):

```bash
python3 SIRCAR/generar_sircar_163_master_dev.py --desde YYYY-MM-DD --hasta YYYY-MM-DD --cuit-agente 30XXXXXXXXX
```

## Notas

- **Jurisdicción (campo 2 y 16)**: para Nakel el default es **`907`**. Si el estudio pide otro (p. ej. `920`), usar `--jurisdiccion 920` (y opcionalmente `--jurisdiccion-sujeto 920`).
- La lista exacta de columnas finales se está alineando con el cuadro del estudio (captura en el chat).
- Por ahora el export prioriza: **Fecha, Cuenta, Débito, Crédito, Etiqueta/Detalle**, que son las columnas visibles en la captura.

