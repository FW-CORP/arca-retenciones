# nakel_scripts

Repositorio interno **FW-CORP / Nakel** para **scripts y utilidades** que no viven dentro de un módulo Odoo: exportaciones fiscales, informes puntuales, conversiones de layout, etc.

Cada bloque de trabajo convive en **su propia carpeta** en la raíz (monorepo simple). Así se evita mezclar dependencias, documentación y convenciones de nombres entre proyectos distintos.

## Contenido actual

| Carpeta | Descripción |
|---------|-------------|
| [`ARCA-RETENCIONES/`](ARCA-RETENCIONES/README.md) | **Retenciones** SICORE/SIRCAR + **percepciones** IIBB (`PERCEIIBB/`), plantillas y documentación en un solo árbol. |

## Cómo usar un subproyecto

1. Entrá a la carpeta del proyecto (ej. `ARCA-RETENCIONES/`).
2. Seguí el `README.md` de esa carpeta (credenciales Odoo suelen salir de `config_nakel` + `.env` fuera del repo).

## Convenciones del repo

- **Sin secretos en el árbol**: no subir `.env`, claves API ni contraseñas; las salidas generadas (`**/out/`) están ignoradas por git.
- **Un README por carpeta** que explique alcance, requisitos y comandos típicos.
- **Sin rutas absolutas fijas** en scripts compartidos: cada subproyecto documenta cómo resolver `PYTHONPATH` / `config_nakel` (ej. `ARCA-RETENCIONES` usa `NAKEL_CONFIG_ROOT` o búsqueda hacia arriba de `config_nakel.py`).

## Clonar

```bash
git clone git@github.com:FW-CORP/nakel_scripts.git
cd nakel_scripts
```
