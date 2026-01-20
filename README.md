# Cotizaciones-Antartida

Este repositorio genera automáticamente archivos Excel de cotización a partir de
una plantilla existente y un archivo JSON con los datos de proyecto y conceptos.
El objetivo es crear una nueva cotización sin editar manualmente la plantilla.

## Estructura del repositorio

- `templates/`: plantillas originales **(no modificar ni renombrar)**.
- `data/`: datos fuente en JSON.
- `scripts/`: scripts de generación.
- `dist/`: archivos generados.
- `.github/workflows/`: automatización con GitHub Actions.

## Cómo editar `data/cotizacion.json`

El archivo `data/cotizacion.json` tiene dos secciones:

### `proyecto`

- `nombre_hoja`: nombre de la hoja generada en el Excel.
- `cliente`: nombre del cliente.
- `direccion`: dirección completa.
- `telefono`: teléfono de contacto.
- `fecha`: fecha de la cotización (formato libre).
- `validez_dias`: número de días de validez.
- `empresa`: nombre de la empresa.

### `conceptos`

Cada elemento del arreglo `conceptos` representa una partida con los campos:

- `descripcion`: texto del concepto (puede incluir saltos de línea).
- `unidad`: unidades o UDM.
- `cantidad`: cantidad.
- `precio_unitario`: precio unitario.

## Cómo se reemplazan los encabezados

El script busca celdas ancla en la hoja para escribir los valores de cabecera:

- Cliente: busca textos como `Cliente` o `Cliente:`.
- Dirección: busca `Dirección` o `Direccion`.
- Teléfono: busca `Teléfono` o `Telefono`.
- Fecha: busca `Fecha del presupuesto` o `Fecha`.
- Validez: busca `Validez`.

**Criterio de escritura:** si la celda a la derecha está vacía, el valor se
escribe allí; si no, se reemplaza el contenido de la celda ancla. Si el texto de
la plantilla no coincide con estos anclajes, ajusta la plantilla o los anclajes
en el script.

## Generación del Excel

Para generar la cotización localmente:

```bash
python scripts/generate_excel.py
```

El script copia la hoja plantilla `Lomas Country Temixco` y crea una nueva hoja
con el nombre definido en `proyecto.nombre_hoja`.

El archivo generado se guarda como `dist/Cotizacion_Generada.xlsx`.

## Descargar desde GitHub Actions

1. Abre la pestaña **Actions** del repositorio.
2. Selecciona el workflow **Build Excel Quote**.
3. En la ejecución más reciente, descarga el artifact llamado
   **cotizacion-excel**, que contiene `dist/Cotizacion_Generada.xlsx`.
