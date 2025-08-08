import {
    INodeType,
    INodeTypeDescription,
    IExecuteFunctions,
    INodeExecutionData,
    IDataObject,
} from 'n8n-workflow';

import { Workbook, Worksheet } from 'exceljs';
import * as fs from 'fs';
import { promises as fsPromises } from 'fs'; // Importar promesas de FS
import * as path from 'path';

// ... (la descripción del nodo y las propiedades no cambian, se omiten por brevedad)
const description: INodeTypeDescription = {
    displayName: 'Sheet Manager',
    name: 'sheetManager',
    icon: 'file:gogh.svg',
    group: ['transform'],
    version: 1,
    subtitle: '={{ $parameter["operation"] + ": " + $parameter["sheetName"] }}',
    description: 'Crea, edita o borra archivos XLSX en tus workflows de n8n',
    defaults: {
        name: 'Sheet Manager',
    },
    inputs: ['main'] as any,
    outputs: ['main'] as any,
    credentials: [],
    properties: [
        // Operación
        {
            displayName: 'Operación',
            name: 'operation',
            type: 'options',
            options: [
                {
                    name: 'Leer Archivo',
                    value: 'readFile',
                },
                { name: 'Visualizar Hoja', value: 'view' },
                { name: 'Crear o Agregar', value: 'create' },
                { name: 'Editar', value: 'edit' },
                { name: 'Borrar', value: 'deleteFile' },
            ],
            default: 'view',
        },

        // Archivo
        {
            displayName: 'Archivo (.xlsx)',
            name: 'filePath',
            type: 'string',
            default: '',
            required: false,
            description: 'Ruta del archivo XLSX. Relativa a /data/sheet-manager si no es absoluta. por defecto se toma `/tmp/data.xlsx`',
        },
        // Nombre de archivo
        {
            displayName: 'Nombre del Archivo',
            name: 'fileName',
            type: 'string',
            default: '',
            required: false,
            description: 'Opcional. Si se especifica, reemplaza el nombre del archivo en la ruta. Por ejemplo, "mi_archivo.xlsx".',
            displayOptions: {
                show: {
                    operation: ['create'],
                },
            },
        },


        // Nombre de hoja
        {
            displayName: 'Hoja',
            name: 'sheetName',
            type: 'string',
            description: 'Nombre de la hoja a procesar',
            displayOptions: {
                show: {
                    operation: ['view', 'create', 'edit'],
                },
            },
            default: '',
        },

        // Modo Append (solo crear)
        {
            displayName: 'Modo Append',
            name: 'append',
            type: 'boolean',
            default: false,
            description: 'Si el archivo existe, agrega los datos al final en lugar de sobreescribir.',
            displayOptions: {
                show: {
                    operation: ['create'],
                },
            },
        },

        // Encabezados manuales (opcional)
        {
            displayName: 'Encabezados personalizados',
            name: 'headers',
            type: 'fixedCollection',
            typeOptions: {
                multipleValues: true,
            },
            default: {}, // Usar objeto vacío por defecto
            placeholder: 'Agregar encabezados manuales',
            description: 'Opcional. Si se omite, se infieren desde el JSON de datos.',
            displayOptions: {
                show: {
                    operation: ['create'],
                },
            },
            options: [
                {
                    displayName: 'Encabezados',
                    name: 'headersValues',
                    values: [
                        {
                            displayName: 'Encabezado',
                            name: 'header',
                            type: 'string',
                            default: '',
                            placeholder: 'Ej. Nombre',
                        },
                    ],
                },
            ],
        },
        {
            displayName: 'Valor por Defecto para Vacíos',
            name: 'defaultFillValue',
            type: 'string',
            default: 'null',
            description: 'Valor a usar cuando una propiedad no está presente en un objeto de datos. Usa "null" para dejar nulo.',
            displayOptions: {
                show: {
                    operation: ['create'],
                },
            },
        },

        // Datos JSON (crear)
        {
            displayName: 'Datos JSON',
            name: 'data',
            type: 'json',
            required: true,
            default: '[]',
            description:
                'Array de objetos para insertar: [{ "col1": "val1" }, ...]',
            displayOptions: {
                show: {
                    operation: ['create'],
                },
            },
        },

        // Propiedades para 'Editar'
        {
            displayName: 'Columna de Condición',
            name: 'conditionColumn',
            type: 'string',
            default: '',
            required: true,
            description: 'Columna para encontrar la fila a modificar',
            displayOptions: { show: { operation: ['edit'] } },
        },
        {
            displayName: 'Valor de Condición',
            name: 'conditionValue',
            type: 'string',
            default: '',
            required: true,
            description: 'Valor exacto a buscar en la columna de condición',
            displayOptions: { show: { operation: ['edit'] } },
        },
        {
            displayName: 'Columna a Modificar',
            name: 'targetColumn',
            type: 'string',
            default: '',
            required: false,
            description: 'Columna donde se aplicará el nuevo valor. Por defecto, es la columna de condición.',
            displayOptions: { show: { operation: ['edit'] } },
        },
        {
            displayName: 'Nuevo Valor',
            name: 'newValue',
            type: 'string',
            default: '',
            required: true,
            description: 'El nuevo valor para la celda',
            displayOptions: { show: { operation: ['edit'] } },
        },
    ],
};

export class SheetManager implements INodeType {
    description = description;


    async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
        /**
    * Extrae los encabezados de una hoja de cálculo de forma robusta.
    */
        function _getHeaders(sheet: Worksheet): string[] {
            if (sheet.columns && sheet.columns.length > 0) {
                const headers = sheet.columns.map(col => (typeof col.header === 'string' ? col.header.trim() : '')).filter(Boolean);
                if (headers.length > 0) return headers;
            }

            const firstRow = sheet.getRow(1);
            if (!firstRow) return [];

            const headers: string[] = [];
            firstRow.eachCell({ includeEmpty: true }, (cell) => {
                const cellValue = cell.value ? String(cell.value).trim() : '';
                if (cellValue) {
                    headers.push(cellValue);
                }
            });
            return headers;
        }

        /**
         * Extrae un conjunto único de todas las claves (encabezados) de un array de objetos.
         */
        function _extractAllHeadersFromJson(data: IDataObject[]): string[] {
            const headersSet = new Set<string>();
            for (const row of data) {
                if (row && typeof row === 'object') {
                    Object.keys(row).forEach(k => headersSet.add(k));
                }
            }
            return Array.from(headersSet);
        }

        const items = this.getInputData();
        const returnData: INodeExecutionData[] = [];

        for (let i = 0; i < items.length; i++) {
            const rawPath = this.getNodeParameter('filePath', i, "/tmp/data.xlsx") as string;
            const sheetName = this.getNodeParameter('sheetName', i, "") as string;
            const operation = this.getNodeParameter('operation', i, "") as string;


            const defaultFillValueRaw = this.getNodeParameter('defaultFillValue', i, 'null') as string;
            const defaultFillValue = defaultFillValueRaw === 'null' ? null : defaultFillValueRaw;

            const filePath = rawPath.startsWith('/') ? rawPath : path.join('/data/sheet-manager', rawPath);
            fs.mkdirSync(path.dirname(filePath), { recursive: true });

            const workbook = new Workbook();
            const fileExists = fs.existsSync(filePath);

            if (fileExists && operation !== 'deleteFile') {
                await workbook.xlsx.readFile(filePath);
            }

            // =================================================================
            // === OPERACIÓN: LEER ARCHIVO (BINARIO)
            // =================================================================
            if (operation === 'readFile') {
                if (!fileExists) {
                    returnData.push({ json: { success: false, message: `El archivo "${path.basename(filePath)}" no existe.` } });
                } else {
                    const buffer = await fsPromises.readFile(filePath);
                    returnData.push({
                        json: { success: true },
                        binary: {
                            file: {
                                data: buffer.toString('base64'),
                                mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                fileName: path.basename(filePath),
                            }
                        }
                    });
                }
                continue;
            }


            // =================================================================
            // === OPERACIÓN: VER
            // =================================================================
            if (operation === 'view') {
                if (!fileExists) throw new Error(`El archivo "${path.basename(filePath)}" no existe.`);
                const sheet = workbook.getWorksheet(sheetName);
                if (!sheet) throw new Error(`La hoja "${sheetName}" no existe en el archivo.`);

                const headers = _getHeaders(sheet);
                const jsonData: IDataObject[] = [];

                if (headers.length > 0) {
                    sheet.eachRow((row, rowNumber) => {
                        if (rowNumber === 1) return; // Omitir fila de encabezados
                        const rowObject: IDataObject = {};
                        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                            const header = headers[colNumber - 1];
                            if (header) {
                                rowObject[header] = cell.value;
                            }
                        });
                        // Solo agregar el objeto si no está completamente vacío
                        if (Object.keys(rowObject).length > 0) {
                            jsonData.push(rowObject);
                        }
                    });
                }

                const buffer = await workbook.xlsx.writeBuffer();

                // CORRECCIÓN: La estructura del objeto binario era incorrecta.
                returnData.push({
                    json: { data: jsonData },
                    binary: { // Las propiedades deben estar directamente aquí
                        file: {
                            data: Buffer.from(buffer).toString('base64'),
                            mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                            fileName: path.basename(filePath),
                        }
                    },
                });
                continue;
            }

            // =================================================================
            // === OPERACIÓN: CREAR
            // =================================================================
            if (operation === 'create') {
                const append = this.getNodeParameter('append', i, false) as boolean;
                const userHeadersRaw = this.getNodeParameter('headers', i) as { headersValues?: { header: string }[] };
                const userHeaders = userHeadersRaw.headersValues?.map(h => h.header).filter(Boolean) ?? [];
                const fileName = this.getNodeParameter('fileName', i, null) as string;

                let data = this.getNodeParameter('data', i, '[]');
                if (typeof data === 'string') {
                    try { data = JSON.parse(data); } catch { throw new Error('El campo "data" no es un JSON válido.'); }
                }
                if (!Array.isArray(data)) throw new Error('El campo "data" debe ser un array de objetos.');

                let sheet = workbook.getWorksheet(sheetName);
                const isNewSheet = !sheet;
                if (isNewSheet) {
                    sheet = workbook.addWorksheet(sheetName);
                }

                const existingHeaders = (fileExists && append && !isNewSheet) ? _getHeaders(sheet as Worksheet) : [];
                const dataHeaders = _extractAllHeadersFromJson(data as IDataObject[]);

                const unifiedHeaders = Array.from(new Set([...existingHeaders, ...userHeaders, ...dataHeaders]));

                if (!append && sheet!.rowCount > 0) {
                    sheet!.spliceRows(1, sheet!.rowCount);
                }

                sheet!.columns = unifiedHeaders.map(header => ({ header, key: header }));

                const normalizedRows = data.map((row: any) => {
                    const normalized: IDataObject = {};
                    for (const header of unifiedHeaders) {
                        normalized[header] = row.hasOwnProperty(header) ? row[header] : defaultFillValue;
                    }
                    return normalized;
                });

                if (normalizedRows.length > 0) {
                    sheet!.addRows(normalizedRows);
                }

                // if (data.length > 0) {
                //     sheet!.addRows(data);
                // }

                workbook.creator = 'Sheet Manager (n8n)';
                workbook.lastModifiedBy = 'n8n workflow';
                workbook.created = new Date();
                workbook.modified = new Date();
                workbook.title = fileName || path.basename(filePath);

                await workbook.xlsx.writeFile(filePath);
                returnData.push({ json: { success: true, message: `Archivo "${path.basename(filePath)}" guardado.` } });
                continue;
            }

            // =================================================================
            // === OPERACIÓN: EDITAR
            // =================================================================
            if (operation === 'edit') {
                if (!fileExists) throw new Error(`El archivo "${path.basename(filePath)}" no existe.`);
                const sheet = workbook.getWorksheet(sheetName);
                if (!sheet) throw new Error(`La hoja "${sheetName}" no existe.`);

                const conditionColumn = this.getNodeParameter('conditionColumn', i) as string;
                const conditionValue = this.getNodeParameter('conditionValue', i, '') as string | number;
                let targetColumn = (this.getNodeParameter('targetColumn', i) as string) || conditionColumn;
                const newValue = this.getNodeParameter('newValue', i) as string | number | null;

                const headers = _getHeaders(sheet);

                // CORRECCIÓN: Búsqueda de columnas insensible a mayúsculas/minúsculas.
                const lowerCaseHeaders = headers.map(h => h.toLowerCase());
                const conditionColIndex = lowerCaseHeaders.indexOf(conditionColumn.toLowerCase()) + 1;
                const targetColIndex = lowerCaseHeaders.indexOf(targetColumn.toLowerCase()) + 1;

                if (conditionColIndex === 0) throw new Error(`La columna de condición "${conditionColumn}" no fue encontrada. Encabezados disponibles: ${headers.join(', ')}`);
                if (targetColIndex === 0) throw new Error(`La columna objetivo "${targetColumn}" no fue encontrada. Encabezados disponibles: ${headers.join(', ')}`);

                let updated = false;
                sheet.eachRow((row, rowNumber) => {
                    if (rowNumber === 1) return; // Omitir encabezado
                    const cellToCompare = row.getCell(conditionColIndex);
                    if (cellToCompare.value !== null && String(cellToCompare.value).trim() === String(conditionValue).trim()) {
                        row.getCell(targetColIndex).value = newValue;
                        updated = true;
                    }
                });

                if (!updated) {
                    returnData.push({ json: { success: false, message: `No se encontró ninguna fila con el valor "${conditionValue}" en la columna "${conditionColumn}".` } });
                } else {
                    await workbook.xlsx.writeFile(filePath);
                    returnData.push({ json: { success: true, message: 'Archivo actualizado correctamente.' } });
                }
                continue;
            }

            // =================================================================
            // === OPERACIÓN: BORRAR
            // =================================================================
            if (operation === 'deleteFile') {
                if (!fileExists) {
                    returnData.push({ json: { success: false, message: 'El archivo no existe, no se puede borrar.' } });
                } else {
                    try {
                        // CORRECCIÓN: Usar la versión asíncrona y manejar errores.
                        await fsPromises.unlink(filePath);
                        returnData.push({ json: { success: true, message: `Archivo "${path.basename(filePath)}" borrado.` } });
                    } catch (error: any) {
                        throw new Error(`No se pudo borrar el archivo. Causa: ${error.message}. Verifique los permisos de la carpeta /data/sheet-manager.`);
                    }
                }
                continue;
            }
        }

        return this.prepareOutputData(returnData);
    }
}