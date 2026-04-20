## ARCA Retenciones – análisis de layouts (TXT)

Parte del monorepo [`nakel_scripts`](../README.md): scripts y documentación para exportes **ARCA / SICORE / SIRCAR** desde Odoo (solo lectura).

Este directorio contiene 3 archivos `.TXT` que se entregan al contador para presentar/armar retenciones (ARCA/DGR/SIRCAR).

### Documentación

- `Documentacion/arca-doc.md`: especificación SICORE 9.0 (posiciones/longitudes/reglas).
- `Documentacion/ARCA_RETENCIONES_LAYOUTS_Y_MAPEO_ODOO.md`: guía unificada (layouts reales + mapeo a Odoo + scripts solo lectura).
- `Documentacion/MANUAL_ARCA_RETENCIONES_QUINCENA.md`: **manual corto** (quincena / rango de fechas, comandos `run_quincena`).

### Uso habitual (certificado / quincena)

Desde esta carpeta:

```bash
python3 SICORE/run_quincena.py --desde YYYY-MM-DD --hasta YYYY-MM-DD
python3 SIRCAR/run_quincena.py --desde YYYY-MM-DD --hasta YYYY-MM-DD --cuit-agente 30XXXXXXXXX
```

Salidas: `SICORE/out/` y `SIRCAR/out/` (nombres con rango de fechas).

### Scripts (Odoo `master_dev`, solo lectura)

Todos leen de `config_nakel.ODOO_CONFIG_MASTER_DEV` salvo anotación contraria.

| Script | Salida típica | Rol |
|--------|----------------|-----|
| `SICORE/run_quincena.py` | `SICORE/out/SICORE_V9_RET_GAN_…TXT` | Atajo: SICORE + validación |
| `SICORE/generar_sicore_v9_retenciones.py` | `SICORE/out/SICORE_V9_RETENCIONES_GANANCIAS.TXT` | SICORE v9 (Ganancias), modo avanzado |
| `SIRCAR/run_quincena.py` | `SIRCAR/out/SIRCAR_163_…TXT` | Atajo: SIRCAR 163 |
| `SIRCAR/generar_sircar_163_master_dev.py` | `SIRCAR/out/SIRCAR_163.TXT` | SIRCAR ancho fijo 163, modo avanzado |
| `SICORE/tools/generar_rgan_cpa_master_dev.py` | `SICORE/out/RGAN_CPA.TXT` | Ganancias ancho fijo 145 (layout estudio) |
| `SICORE/tools/generar_op_odoo_master_dev.py` | `SICORE/out/OP_odoo.xlsx` | Órdenes de pago (planilla) |
| `SICORE/tools/generar_ret_gan_mayor_odoo_master_dev.py` | `SICORE/out/RET_GAN_odoo.xlsx` | Mayor ret. Ganancias |
| `SIRCAR/tools/generar_sircar_mayor_odoo_master_dev.py` | `SIRCAR/out/SIRCAR_mayor.csv` | Mayor / extracto SIRCAR (CSV) |
| `SIRCAR/tools/generar_ret_iibb_mayor_odoo_master_dev.py` | `SIRCAR/out/RET_IIBB_odoo.xlsx` | Mayor ret. IIBB/SIRCAR |
| `SIRCAR/tools/generar_ret_dgr_master_dev.py` | `SIRCAR/out/RET-DGR.TXT` (CSV tipo SIRCAR) | DGR/IIBB en CSV (layout estudio) |
| `SIRCAR/tools/generar_ret_dgr_ancho_fijo_master_dev.py` | `SIRCAR/out/RET-DGR.TXT` (ancho fijo) | DGR layout contador |

### Archivos

- `RET-DGR-SIRCAR.TXT`
  - **Formato**: CSV (separado por comas), 1 registro por retención.
  - **Ejemplo**:
    - `00001,1,1,000000018150,30707971297  ,02/03/2026,3043512.28, 3.75,114131.71,001,907`
  - **Campos inferidos (por posición)**:
    - `col1` **nro_registro**: correlativo (5 dígitos, con ceros a la izquierda).
    - `col2` **lote**: constante `1` (parece “lote/presentación”).
    - `col3` **sub_lote**: constante `1`.
    - `col4` **nro_orden_pago**: `000000018150` (correlativo interno; coincide con los otros archivos).
    - `col5` **cuit_sujeto**: 11 dígitos (puede venir con espacios al final).
    - `col6` **fecha_retencion**: `dd/mm/yyyy`.
    - `col7` **base_imponible**: decimal con punto (`.`).
    - `col8` **alicuota**: porcentaje (ej. `3.75`).
    - `col9` **importe_retenido**: decimal con punto (`.`).
    - `col10` **codigo_regimen?**: `001` (constante en la muestra).
    - `col11` **codigo_jurisdiccion/agente?**: `907` (constante en la muestra).

- `RET-DGR.TXT`
  - **Formato**: ancho fijo (sin separadores), con 1 **encabezado** + N **detalles**.
  - **Pistas fuertes**:
    - Encabezado con **período** y **monto en letras** (ej. `SON PESOS ...`) + datos del agente.
    - Detalles con importes “serializados” (probable *100), CUIT con guiones y datos del proveedor (razón social/domicilio/localidad/CP/provincia).

- `RGAN_CPA.TXT`
  - **Formato**: ancho fijo (sin separadores), 1 registro por retención.
  - **Pistas fuertes**:
    - Arranca con `06` → probable **tipo de registro**.
    - Repite **fecha** varias veces en el mismo registro.
    - CUIT aparece como **`80` + `CUIT(11)`** (13 dígitos), p.ej. `8030707971297`.
    - Montos usan coma decimal (`,`) y padding con espacios.

### Qué hay que extraer de Odoo (hipótesis operacional)

Para construir estos archivos desde Odoo, el “origen” suele ser **el pago a proveedor** (orden de pago) y las **líneas de retención** generadas por impuestos de retención:

- **Proveedor**: CUIT (`res.partner.vat`).
- **Orden/pago**: número interno (ej. `000000018150`) y fecha de pago.
- **Base imponible**: base sobre la que aplica la retención (según configuración del impuesto).
- **Alícuota**: % de retención.
- **Importe retenido**: monto retenido.
- **Jurisdicción/régimen**: códigos (ej. `010`, `001`, `907`, `02170781`) que suelen venir de la parametrización fiscal (retención IIBB / SIRCAR / Ganancias).

### Pendiente para cerrar layout 100%

- Confirmar **longitudes exactas** (posiciones) de `RGAN_CPA.TXT` y `RET-DGR-SIRCAR.TXT` comparando con especificación del organismo/aplicativo o con un export “oficial” del contador.
- Identificar con certeza qué representan los códigos constantes `001` y `907` y el bloque `02170781`.

### Seguridad y qué no va al repositorio

- En este árbol **no** se versionan credenciales: los scripts leen `config_nakel` (fuera de esta carpeta) y/o variables de entorno (`nakel/.env`, etc.).
- Las carpetas `**/out/` están en `.gitignore`: ahí suelen quedar **TXT/CSV/XLSX generados** con datos reales de Odoo (CUIT, montos). No los subas a un repo público sin revisar.
- Los `.TXT` de muestra en la raíz / `Documentacion/` son referencia del estudio; pueden contener **CUIT de ejemplo** (no son contraseñas, pero sí datos fiscales de muestra).

