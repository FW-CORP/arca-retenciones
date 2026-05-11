---
title: exportador-excel — planillas para estudio
updated: 2026-05-11
---

## Objetivo

Generar planillas **XLSX** (vía **CSV + LibreOffice headless**) con el detalle de:

1) **Retenciones aplicadas** (SIRCAR / IIBB) — pagos a **proveedores** (agente de retención).
2) **IIBB sufrida** en compras — facturas de **proveedor** (`in_invoice` / `in_refund`).
3) **Percepciones IIBB en ventas** — misma grilla que suele pedir el estudio (`Nro.Inscripcion` … `Nro.Percepcion`). En la última columna se informa el **número de comprobante fiscal** (solo dígitos o `PV-NNNNNNNN`), no el nombre del impuesto ni la constancia inexistente en Odoo.
4) **Percepción TCI multilateral en ventas** — misma grilla; filtra solo impuestos de venta cuyo nombre contiene por defecto **Percepcion MCR TCI Multilateral** (script aparte: `percepciones_tci_ventas_estudio.py`).

Fuente de datos: Odoo `master_dev` por XML-RPC (solo lectura), usando la misma base (`tax_base_amount`) y alícuota (`account.tax.amount`) que alimentan `RET-DGR-SIRCAR.TXT` donde aplica.

## Requisitos

- `libreoffice` disponible en el sistema (se usa `--headless`).
- `config_nakel.py` accesible (ver `ARCA-RETENCIONES/README.md` y `NAKEL_CONFIG_ROOT`).

## Uso

Desde `ARCA-RETENCIONES/`:

```bash
python3 exportador-excel/retenciones_aplicadas_sircar_iibb.py --desde 2026-04-01 --hasta 2026-04-15
python3 exportador-excel/retenciones_sufridas_iibb.py --desde 2026-04-01 --hasta 2026-04-15
python3 exportador-excel/retenciones_ganancias_rgan_cpa.py --desde 2026-04-01 --hasta 2026-04-15
python3 exportador-excel/percepciones_iibb_ventas_estudio.py --desde 2026-04-01 --hasta 2026-04-30 --layout-estudio
# Opcional: columna Nro.Percepcion como PV-número (estilo 0050-00000014)
python3 exportador-excel/percepciones_iibb_ventas_estudio.py --desde 2026-04-01 --hasta 2026-04-30 --nro-factura-formato pv_guion_num
python3 exportador-excel/percepciones_tci_ventas_estudio.py --desde 2026-04-01 --hasta 2026-04-30 --layout-estudio
# Incluir también «TCI sin Constancia» u otro nombre: ajustar --impuesto-nombre-contiene "Percepcion MCR TCI"
```

Salidas: `exportador-excel/out/*.xlsx`

