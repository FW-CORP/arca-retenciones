---
title: exportador-excel — planillas para estudio
updated: 2026-04-21
---

## Objetivo

Generar planillas **XLSX** (vía **CSV + LibreOffice headless**) con el detalle de:

1) **Retenciones aplicadas** (SIRCAR / IIBB) — pagos a **proveedores** (agente de retención).
2) **IIBB sufrida** en compras — facturas de **proveedor** (`in_invoice` / `in_refund`).

Fuente de datos: Odoo `master_dev` por XML-RPC (solo lectura), usando la misma base (`tax_base_amount`) y alícuota (`account.tax.amount`) que alimentan `RET-DGR-SIRCAR.TXT`.

## Requisitos

- `libreoffice` disponible en el sistema (se usa `--headless`).
- `config_nakel.py` accesible (ver `ARCA-RETENCIONES/README.md` y `NAKEL_CONFIG_ROOT`).

## Uso

Desde `ARCA-RETENCIONES/`:

```bash
python3 exportador-excel/retenciones_aplicadas_sircar_iibb.py --desde 2026-04-01 --hasta 2026-04-15
python3 exportador-excel/retenciones_sufridas_iibb.py --desde 2026-04-01 --hasta 2026-04-15
python3 exportador-excel/retenciones_ganancias_rgan_cpa.py --desde 2026-04-01 --hasta 2026-04-15
```

Salidas: `exportador-excel/out/*.xlsx`

