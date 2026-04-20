---

## SICORE 9.0 — Diseño de Registro TXT para Importación de Retenciones

### 1. Consideraciones generales del archivo

- Codificación: ASCII / ANSI
- Separador de registros: retorno de carro (CR+LF) al final de cada línea
- *Separador decimal: coma (,)* — esto se configura en el aplicativo en Importar/Exportar Retenciones → Configuración de Importación de Retenciones
- *Todos los registros deben tener la misma longitud* (error clásico si hay diferencias)
- Alineación alfanumérica: *izquierda*, relleno con espacios
- Alineación numérica: *derecha*, relleno con ceros
- Formato de fecha: dd/mm/aaaa
- CUIT/CUIL/CDI: *sin guiones*
- Números de comprobante: *sin separadores* (ej.: 000112334599, no 0001-12334599)

---

### 2. Cambio crítico en versión 9 vs. versión 8

La versión 9.0 introdujo *cambios en el diseño de registro* respecto a la v8. En particular:

- El *código de impuesto pasa a 4 dígitos* (en v8 era 3). Ejemplo: 217 → 0217.
- Se incorporó el campo obligatorio *CONDICIÓN* (posiciones 77-78 del registro estándar).
- Foros de práctica confirmaron que el código de impuesto debe tener 4 dígitos y que las posiciones 77-78 del campo CONDICIÓN deben completarse correctamente (p.ej. 00 cuando no aplica condición específica), ya que errores en ese campo provocan el rechazo del registro. [Ignacio online](https://www.ignacioonline.com.ar/aplicativos-nueva-version-sicore-version-9/)
- Con el release 18 (octubre 2024), ARCA modificó nuevamente el separador de campos, lo que genera errores de importación en archivos generados con el formato anterior. [Contabilium AR](https://ayuda.contabilium.com/hc/es/articles/37128582109203-Tengo-errores-al-importar-mi-archivo-txt-en-el-aplicativo-Sicore)

---

### 3. Diseño de registro estándar — Retenciones (sin Beneficiarios del Exterior)

El formato estándar que el propio SICORE expone bajo Configuración de Importación de Retenciones es el siguiente:

| Pos. Inicio | Longitud | Campo | Tipo | Observaciones |
|-------------|----------|-------|------|---------------|
| 1 | 2 | Código de comprobante | AN | Ver Tabla A |
| 3 | 10 | Fecha del comprobante | AN | dd/mm/aaaa |
| 13 | 16 | Número del comprobante | AN | Sin separadores; para C.1116 formato especial |
| 29 | 16 | Importe del comprobante | N | Con coma decimal; ≥ 0 |
| 45 | 4 | Código de impuesto | N | *4 dígitos desde v9* (ej. 0217) |
| 49 | 4 | Código de régimen | N | 4 dígitos; ver Tabla B |
| 53 | 1 | Código de operación | N | Ver Tabla C |
| 54 | 14 | Base de cálculo | N | Con coma decimal |
| 68 | 10 | Fecha de retención | AN | dd/mm/aaaa |
| 78 | 2 | Código de condición | N | *Nuevo campo v9*; ver Tabla D |
| 80 | 1 | Retención a sujeto suspendido | N | Ver Tabla E |
| 81 | 14 | Importe de la retención | N | Con coma decimal; ≥ 0 |
| 95 | 6 | Porcentaje de exclusión | N | 000000 si no aplica |
| 101 | 10 | Fecha del boletín oficial | AN | dd/mm/aaaa; `          ` si no aplica |
| 111 | 2 | Tipo de documento retenido | AN | Ver Tabla F |
| 113 | 20 | Nro. de documento retenido | AN | Sin guiones |

*Longitud total del registro: 132 posiciones* (más CR+LF)

> ⚠️ Para Notas de Crédito se deben agregar 16 posiciones adicionales para el Nro. de Certificado Original.

---

### 4. Tablas de códigos aplicables a Ganancias

*Tabla A — Código de comprobante*

| Código | Descripción |
|--------|-------------|
| 01 | Factura |
| 02 | Recibo |
| 03 | Nota de Crédito |
| 04 | Nota de Débito |
| 05 | Otro comprobante |
| 06 | Orden de Pago |
| 07 | Recibo de Sueldo (solo régimen 160) |
| 08 | Recibo de Sueldo – Devolución (solo régimen 160) |
| 09 | Escritura Pública |
| 10 | C.1116 (regímenes 680, 681, 781–796, entre otros) |
| 11 | Factura 16 dígitos |

*Tabla B — Principales códigos de impuesto Ganancias (4 dígitos)*

| Código | Descripción |
|--------|-------------|
| 0217 | Impuesto a las Ganancias |
| 0218 | Ganancias – Beneficiarios del Exterior |
| 0210 | Ganancias – Régimen Especial RG 830 (usa regímenes del 0217) |

*Tabla C — Código de operación*

| Código | Descripción |
|--------|-------------|
| 1 | Retención |
| 2 | Percepción |
| 4 | Imposibilidad de Retención |

*Tabla D — Código de condición (campo nuevo v9, posiciones 77-78)*

| Código | Descripción |
|--------|-------------|
| 00 | Ninguna |
| 01 | Inscripto |
| 02 | No inscripto |
| 03 | No categorizado |
| 06 | Contratación Hora/Día/Estadía |
| 07 | Contratación Mensual |
| 10 | Inscripto demás sujetos |
| 13 | Venta cosas muebles y locación – alícuota general |
| 14 | Venta cosas muebles y locación – alícuota reducida |
| 15 | Retención sustitutiva |
| 16 | Sujeto suspendido art. 40 inc. A |
| 17 | Sujeto suspendido art. 40 inc. B |
| 18 | Aplica Convenio Doble Imposición |
| 19 | No aplica Convenio Doble Imposición |

> El código de condición válido *depende del régimen. Verificar combinación en el aplicativo: *Consulta de DDJJ → Consulta de Regímenes.

*Tabla E — Retención a sujeto suspendido*

| Código | Descripción |
|--------|-------------|
| 0 | Ninguno |
| 1 | Art. 40 inciso A |
| 2 | Art. 40 inciso B |

*Tabla F — Tipo de documento del retenido*

| Código | Descripción |
|--------|-------------|
| 80 | CUIT |
| 86 | CUIL |
| 87 | CDI |
| 83 | Identificación tributaria del exterior |
| 84 | Documento del exterior |

---

### 5. Principales regímenes de Ganancias (código impuesto 0217)

Los regímenes más habituales en la práctica (4ta categoría, servicios profesionales, etc.) que el técnico necesita conocer para las validaciones:

| Régimen | Descripción |
|---------|-------------|
| 160 | Rentas del trabajo personal en relación de dependencia |
| 202 | Honorarios a directores, síndicos, consejo de vigilancia |
| 207 | Servicios prestados por empresas |
| 212 | Alquileres o arrendamientos de bienes inmuebles |
| 217 | Alquileres o arrendamientos de bienes muebles |
| 219 | Locación de obras y/o servicios |

Para la lista completa vigente consultar dentro del aplicativo: Consulta de DDJJ → Descripción de Regímenes (se actualiza con cada release).

---

### 6. Advertencias críticas para el desarrollo

1. *Campo CONDICIÓN obligatorio desde v9: sin él o con valor inválido para el régimen, el sistema rechaza el registro con error *"la relación entre código de régimen, código de operación y código de condición no es válida".
2. *Código de impuesto en 4 dígitos*: 0217, no 217.
3. *Separador decimal: coma*. Configurar en el aplicativo antes de la primera importación y verificar que no se resetee tras actualizaciones.
4. *Desmarcar "Incluye Operaciones con Beneficiarios del Exterior"* si no se importan ese tipo de operaciones — el aplicativo lo tilda por defecto al instalar/actualizar.
5. Los errores de importación se graban en errimpret.log en el directorio de instalación del SICORE, archivo que puede abrirse con cualquier editor de texto.
6. El diseño de registro exacto y actualizado a r22 se puede verificar directamente en el aplicativo instalado: Importar/Exportar Retenciones → Configuración de Importación de Retenciones → Detallar, donde se lista cada campo con posición y longitud.

---

*Recomendación para el técnico*: la fuente definitiva de verdad son los archivos de diseño de registro que el propio aplicativo expone en la pantalla de configuración de importación, ya que cada release puede ajustar alguna posición o longitud. Conviene exportar un registro de muestra desde el SICORE y compararlo contra lo que genera el sistema contable antes de pasar a producción.