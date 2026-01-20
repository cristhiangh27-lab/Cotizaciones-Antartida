# Cotizaciones-Antartida

Este repositorio genera automáticamente archivos Excel de cotización a partir de
plantillas existentes y un catálogo de conceptos en JSON. Su objetivo es
estandarizar la preparación de cotizaciones sin editar manualmente la plantilla.

## Estructura del repositorio

- `templates/`: plantillas originales **(no modificar ni renombrar)**.
- `data/`: datos fuente en JSON.
- `scripts/`: scripts de generación.
- `dist/`: archivos generados.
- `.github/workflows/`: automatización con GitHub Actions.

## Cómo editar el catálogo de conceptos

Edita `data/catalogo_conceptos.json` y ajusta el arreglo `conceptos`. Cada
concepto utiliza los siguientes campos:

- `partida`: número de partida.
- `clave`: identificador interno.
- `concepto`: nombre corto del concepto.
- `descripcion`: detalle del concepto.
- `unidad`: unidad de medida.
- `cantidad`: cantidad numérica.
- `precio_unitario`: precio unitario.
- `importe`: total de la partida (si la plantilla ya tiene fórmula, se respeta).

El script buscará automáticamente la fila de encabezados (por ejemplo, columnas
con títulos como "concepto", "cantidad" o "importe") y comenzará a insertar los
conceptos en la siguiente fila. Si la plantilla no contiene encabezados claros o
usa nombres distintos, actualiza los textos en la plantilla para que coincidan
con los encabezados esperados o ajusta el script.

## Generación del Excel

Para generar la cotización localmente:

```bash
python scripts/generate_excel.py
```

El archivo generado se guarda como `dist/Cotizacion_Antartida.xlsx`.

## Descargar desde GitHub Actions

1. Abre la pestaña **Actions** del repositorio.
2. Selecciona el workflow **Build Excel Quote**.
3. En la ejecución más reciente, descarga el artifact llamado
   **cotizacion-excel**, que contiene `dist/Cotizacion_Antartida.xlsx`.
