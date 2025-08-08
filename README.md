
# Sheet Manager (n8n Custom Node)

**Sheet Manager** es un nodo personalizado para n8n que permite crear, leer, editar, visualizar y borrar archivos `.xlsx` directamente desde tus flujos de trabajo. Está construido usando `exceljs`, funciona sin credenciales y opera sobre rutas relativas a un directorio de trabajo (`/data/sheet-manager` por defecto).

## Características principales

* Crear archivos XLSX con encabezados personalizados y filas dinámicas.
* Agregar datos a hojas existentes sin sobrescribir (modo append).
* Visualizar contenido de hojas específicas como JSON y archivo binario.
* Editar filas específicas en función de una condición.
* Eliminar archivos XLSX fácilmente.
* Leer archivo completo como binario para descargar o reenviar.

## Instalación

Este nodo es parte de un desarrollo personalizado. Asegúrate de colocarlo en tu directorio de nodos personalizados en n8n:

```bash
~/.n8n/custom-nodes/
```

## Parámetros del Nodo

| Campo                          | Descripción                                                               | Operaciones disponibles |
| ------------------------------ | -------------------------------------------------------------------------- | ----------------------- |
| Operación                     | Define la acción a realizar (view, create, edit, etc)                     | Todos                   |
| Archivo (.xlsx)                | Ruta al archivo. Si no es absoluta, se asume dentro de /data/sheet-manager | Todos                   |
| Nombre del Archivo             | Reemplaza el nombre del archivo en la ruta                                 | create                  |
| Hoja                           | Nombre de la hoja a procesar                                               | view, create, edit      |
| Modo Append                    | Agrega datos al final en lugar de sobrescribir                             | create                  |
| Encabezados personalizados     | Define manualmente los encabezados de columna                              | create                  |
| Valor por Defecto para Vacíos | Se usa cuando falta una propiedad en un objeto                             | create                  |
| Datos JSON                     | Array de objetos que representan filas                                     | create                  |
| Columna de Condición          | Columna para buscar la fila a modificar                                    | edit                    |
| Valor de Condición            | Valor exacto a buscar en la columna de condición                          | edit                    |
| Columna a Modificar            | Columna a cambiar (si se omite, se usa la de condición)                   | edit                    |
| Nuevo Valor                    | Valor nuevo a asignar                                                      | edit                    |

```json
{
  "operation": "create",
  "filePath": "reporte.xlsx",
  "sheetName": "Ventas",
  "append": true,
  "headers": {
    "headersValues": [{ "header": "Producto" }, { "header": "Cantidad" }]
  },
  "defaultFillValue": "N/A",
  "data": [
    { "Producto": "Manzana", "Cantidad": 10 },
    { "Producto": "Pera" }
  ]
}

```


## Consideraciones

* Si una hoja no existe en `create`, se crea automáticamente.
* En `edit`, la búsqueda es exacta (pero case-insensitive).
* El orden de las columnas no afecta el funcionamiento.
* Las hojas deben tener encabezados en la primera fila.

## Requisitos

* Node.js >= 18
* n8n instalado localmente o en entorno controlado
* Dependencia: `exceljs`

## Autor

Desarrollado por **Elimeleth**

Correo: `elimeleth.contacto@gmail.com`

Optimizado para flujos de trabajo robustos y dinámicos.
