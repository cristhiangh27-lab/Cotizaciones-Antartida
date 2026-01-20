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

- `folio`: folio usado para nombrar la hoja y el archivo generado.
- `cliente`: nombre del cliente.
- `direccion`: dirección completa.
- `telefono`: teléfono de contacto.
- `fecha`: fecha de la cotización (formato libre).
- `validez_dias`: número de días de validez (se mantiene el texto `x 30 dias` en la plantilla).

### `conceptos`

Cada elemento del arreglo `conceptos` representa una partida con los campos:

- `descripcion`: texto del concepto (puede incluir saltos de línea).
- `unidades`: unidades o UDM.
- `precio_unitario`: precio unitario.

## Cómo se reemplazan los encabezados

El script busca celdas ancla en la hoja para escribir los valores de cabecera:

- Cliente: reemplaza la celda que contiene exactamente `Cliente:` por
  `Cliente: {cliente}`.
- Dirección: reemplaza la celda que contiene exactamente `Dirección:` por
  `Dirección: {direccion}`.
- Teléfono: reemplaza la celda que contiene exactamente `Teléfono:` por
  `Teléfono: {telefono}`.
- Fecha: busca `Fecha del presupuesto` y escribe la fecha en la celda contigua
  derecha.
- Título: reemplaza la celda que empieza con `Presupuesto` por
  `Presupuesto {folio}`.

Si el texto de la plantilla no coincide con estos valores exactos, ajusta los
textos de anclaje en el script.

## Generación del Excel

Para generar la cotización localmente:

```bash
python scripts/generate_excel.py
```

El script copia la hoja plantilla `Lomas Country Temixco`, crea una nueva hoja
con el nombre `proyecto.folio` y elimina la hoja plantilla en el archivo
generado.

El archivo generado se guarda como `dist/Cotizacion_{folio}.xlsx`.

## Descargar desde GitHub Actions

1. Abre la pestaña **Actions** del repositorio.
2. Selecciona el workflow **Build Excel Quote**.
3. En la ejecución más reciente, descarga el artifact llamado
   **cotizacion-excel**, que contiene `dist/Cotizacion_{folio}.xlsx`.
