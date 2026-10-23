# Inventario General Eaton - versión final con control Activo

## Página pública
Muestra, al consultar un E-Number activo:

- N.º de pareja
- Mesa
- Listado
- Ubicación a contar
- Horario de colación
- Horario de almuerzo

La página reproduce el video `inventario-anual-bodega.mp4`, que debe permanecer en la raíz del repositorio.

## Datos ficticios de prueba

- `E100001`
- `E100002`
- `E100003`
- `E100004`

## Archivo privado

Usa `CONSULTA INVENTARIO.xlsx` en tu computador. Debe conservar:

- Hoja: `LISTADO`
- Tabla: `DETALLE`

Columnas:

1. `ENUMBER`
2. `N° PAREJA`
3. `MESA`
4. `LISTADO`
5. `UBICACIÓN A CONTAR`
6. `HORARIO DE COLACIÓN`
7. `HORARIO DE ALMUERZO`
8. `ACTIVO`

En `ACTIVO`, usa `SI` para publicar y `NO` para retirar temporalmente a una persona.

## Actualizar la página

Guarda y cierra el Excel. Abre PowerShell en la carpeta y ejecuta:

```powershell
powershell -ExecutionPolicy Bypass -File .ctualizar-datos.ps1
```

El script crea un `asignaciones.json` nuevo. Sube únicamente ese JSON a GitHub.

**Nunca subas el Excel real con E-Numbers al repositorio público.**
