# Actualización del inventario de impresoras

Genera `DATOS_BASE_IMPRESORAS_ACTUALIZADO.xlsx` actualizando el inventario oficial
(`datos base impresoras.xlsx`) con la información de `datos recibidos.xlsx`,
comparando por la columna **CODIGO**.

## Qué hace
- Compara por **CODIGO** normalizando mayúsculas/tildes/espacios y notación científica.
- Para cada código coincidente, actualiza Ciudad, Estado, Supervisor, Usuario, Proyecto y
  Observación con el valor más reciente de DATOS RECIBIDOS (solo cuando difiere y no está vacío).
- Conserva intactos: diseño, formatos, colores, fórmulas, dibujos, nombres de hoja y la
  columna histórica `OBSERVACION 19/05/2026`.
- Crea la hoja **REVISION_MANUAL** con los registros sin coincidencia:
  `NUEVO REGISTRO` (en recibidos, no en base) y `NO ENCONTRADO EN ACTUALIZACION` (en base, no en recibidos).
- Crea la hoja **RESUMEN** con los indicadores de coincidencias y actualizaciones.
- `coincide`/`coindice` se tratan como marcador de validación (no se escriben como observación).

## Ejecutar
```bash
python _tools/build.py     # genera el .xlsx final
python _tools/verify.py    # valida integridad (opcional)
```
