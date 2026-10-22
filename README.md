# Inventario General - Consulta pública

Página estática para consultar pareja, mesa, listado y ubicación mediante E-Number.

## Datos de demostración

- `E100001` → Pareja 7A, Mesa 1, Listado 7, AI03
- `E100002` → Pareja 7B, Mesa 1, Listado 7, AI03
- `E100003` → Pareja 1A, Mesa 3, Listado 47, RA11
- `E100004` → Pareja 12A, Mesa 1, Listado 13, AS01

## Generar datos reales

1. Mantener el Excel fuera del repositorio.
2. Confirmar hoja `LISTADO` y tabla `DETALLE`.
3. Confirmar columnas: `E-NUMBER`, `PAREJA`, `MESA`, `LISTADO`, `UBICACION`, `PRIORIDAD`.
4. Abrir PowerShell en la carpeta del Excel y ejecutar:

```powershell
powershell -ExecutionPolicy Bypass -File .ctualizar-datos.ps1
```

5. Subir únicamente el nuevo `asignaciones.json` al repositorio.

El script selecciona, por cada E-Number, la fila con prioridad numérica más baja. Excluye prioridad `x`, filas sin mesa y registros cuyo E-Number continúe como `ENCRIPTADO`.

## GitHub Pages

Publicar desde `main` y `/(root)`. El archivo de entrada es `index.html`.
