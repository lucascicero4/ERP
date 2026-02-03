/**
 * ERP LUCAS v5.3 - Google Apps Script
 * 
 * CORRECCIONES v5.3:
 * 1. Patrimonio: Nueva hoja dedicada "Patrimonio_App" para lectura/escritura consistente
 * 2. Cuotas: Agregado idGrupo para agrupar cuotas de la misma compra
 * 3. Columna L (12) para idGrupo en la hoja de gastos
 * 
 * INSTRUCCIONES:
 * 1. Copiar este código en Google Apps Script
 * 2. Ejecutar setupPatrimonioSheet() UNA VEZ para crear la hoja de patrimonio
 * 3. Ejecutar setupTrigger() UNA VEZ para instalar el trigger de cuotas
 * 4. Implementar > Nueva implementación > Aplicación web
 * 5. Acceso: "Cualquier persona"
 * 6. Copiar la URL y pegarla en CONFIG.SCRIPT_URL del index.html
 */

// ============================================================
// CONFIGURACIÓN DE HOJAS
// ============================================================
const SHEETS = {
    GASTOS: 'Formulario Gastos',
    PATRIMONIO_APP: 'Patrimonio_App',  // Nueva hoja dedicada para la app
    PATRIMONIO: 'Patrimonio',          // Hoja legacy (backup)
    INVERSIONES: 'Inversiones',
    MOVIMIENTOS: 'Movimientos',
    INGRESOS: 'Ingresos Mensuales',
    CONFIG: 'Config'
};

// Mapeo de columnas del Formulario de Gastos
const FORM_COLUMNS = {
    TIMESTAMP: 0,
    FECHA: 1,
    CATEGORIA: 2,
    MONTO: 3,
    MONEDA: 4,
    MEDIO_PAGO: 5,
    CUOTAS: 6,
    DESCRIPCION: 7
    // Columnas adicionales agregadas por el trigger:
    // 8 (I): Cuota Actual
    // 9 (J): Total Cuotas
    // 10 (K): Mes Pago
    // 11 (L): ID Grupo (NUEVO en v5.3)
};

// ============================================================
// SETUP INICIAL - EJECUTAR UNA VEZ
// ============================================================

/**
 * Crear hoja de patrimonio dedicada para la app
 * EJECUTAR UNA VEZ después de instalar el script
 */
function setupPatrimonioSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEETS.PATRIMONIO_APP);
    
    if (!sheet) {
        sheet = ss.insertSheet(SHEETS.PATRIMONIO_APP);
        
        // Estructura: Cuenta | Saldo USD | Última Actualización
        sheet.getRange('A1:C1').setValues([['Cuenta', 'Saldo USD', 'Última Actualización']]);
        sheet.getRange('A2:C4').setValues([
            ['BBVA', 0, new Date()],
            ['Caja Seguridad', 0, new Date()],
            ['Efectivo', 0, new Date()]
        ]);
        
        // Formato
        sheet.setColumnWidth(1, 150);
        sheet.setColumnWidth(2, 120);
        sheet.setColumnWidth(3, 180);
        sheet.getRange('B2:B4').setNumberFormat('$#,##0.00');
        
        return 'Hoja Patrimonio_App creada exitosamente. Configurá los saldos iniciales manualmente.';
    }
    
    return 'La hoja Patrimonio_App ya existe.';
}

/**
 * Agregar columna idGrupo a la hoja de gastos si no existe
 */
function setupGastosColumns() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.GASTOS);
    
    if (!sheet) return 'Hoja de gastos no encontrada';
    
    // Verificar si ya tiene las columnas
    const lastCol = sheet.getLastColumn();
    
    if (lastCol < 12) {
        // Agregar headers faltantes
        if (lastCol < 9) sheet.getRange(1, 9).setValue('Cuota');
        if (lastCol < 10) sheet.getRange(1, 10).setValue('Total Cuotas');
        if (lastCol < 11) sheet.getRange(1, 11).setValue('Mes Pago');
        if (lastCol < 12) sheet.getRange(1, 12).setValue('ID Grupo');
    }
    
    return 'Columnas configuradas correctamente';
}

/**
 * Instalar trigger para procesar cuotas automáticamente
 */
function setupTrigger() {
    // Eliminar triggers existentes
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(t => {
        if (t.getHandlerFunction() === 'onFormSubmit') {
            ScriptApp.deleteTrigger(t);
        }
    });
    
    // Crear nuevo trigger
    ScriptApp.newTrigger('onFormSubmit')
        .forSpreadsheet(SpreadsheetApp.getActive())
        .onFormSubmit()
        .create();
    
    // También configurar columnas
    setupGastosColumns();
    
    return 'Trigger instalado y columnas configuradas correctamente';
}

// ============================================================
// API ENDPOINTS
// ============================================================
function doGet(e) {
    try {
        const action = e.parameter.action || 'status';
        
        // Para acciones que modifican datos, leer de URL
        if (e.parameter.data) {
            const data = JSON.parse(decodeURIComponent(e.parameter.data));
            return handleAction(action, data);
        }
        
        let result;
        switch(action) {
            case 'getAll':
                result = getAllData();
                break;
            case 'getExpenses':
                result = getExpenses();
                break;
            case 'getPatrimonio':
                result = getPatrimonio();
                break;
            case 'getInversiones':
                result = getInversiones();
                break;
            case 'getMovimientos':
                result = getMovimientos();
                break;
            case 'getIngresos':
                result = getIngresos();
                break;
            case 'test':
                result = { 
                    success: true, 
                    message: 'API v5.3 funcionando', 
                    timestamp: new Date().toISOString(),
                    hojaGastos: SHEETS.GASTOS,
                    hojaPatrimonio: SHEETS.PATRIMONIO_APP
                };
                break;
            case 'debug':
                result = debugSheet();
                break;
            case 'setupPatrimonio':
                result = { success: true, message: setupPatrimonioSheet() };
                break;
            case 'setupTrigger':
                result = { success: true, message: setupTrigger() };
                break;
            default:
                result = { success: true, message: 'ERP Lucas API v5.3', status: 'online' };
        }
        
        return jsonResponse(result);
    } catch(error) {
        console.error('Error en doGet:', error);
        return jsonResponse({ success: false, error: error.toString(), stack: error.stack });
    }
}

function doPost(e) {
    try {
        const data = JSON.parse(e.postData.contents);
        return handleAction(data.action, data.data || data);
    } catch(error) {
        console.error('Error en doPost:', error);
        return jsonResponse({ success: false, error: error.toString() });
    }
}

function handleAction(action, data) {
    let result;
    
    console.log('Action:', action, 'Data:', JSON.stringify(data));
    
    switch(action) {
        case 'deleteExpense':
            result = deleteExpenseByRow(data.rowIndex);
            break;
        case 'updatePatrimonio':
            result = updatePatrimonio(data);
            break;
        case 'addInversion':
            result = addInversion(data);
            break;
        case 'updateInversion':
            result = updateInversion(data);
            break;
        case 'deleteInversion':
            result = deleteInversion(data.id);
            break;
        case 'addMovimiento':
            result = addMovimiento(data);
            break;
        case 'setMonthlyIncome':
            result = setMonthlyIncome(data);
            break;
        case 'setDolaresAhorro':
            result = setDolaresAhorro(data.amount);
            break;
        case 'addExpense':
            result = addExpense(data);
            break;
        case 'addExpenseWithCuotas':
            result = addExpenseWithCuotas(data.expense, data.cuotas, data.mesesPago);
            break;
        case 'updateExpense':
            result = updateExpense(data);
            break;
        default:
            result = { success: false, error: 'Acción no reconocida: ' + action };
    }
    
    return jsonResponse(result);
}

function jsonResponse(data) {
    return ContentService
        .createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// LECTURA DE DATOS
// ============================================================
function getAllData() {
    try {
        return {
            success: true,
            expenses: getExpenses().data || [],
            patrimonio: getPatrimonio().data || { bbva: 0, caja: 0, efectivo: 0 },
            inversiones: getInversiones().data || [],
            movimientos: getMovimientos().data || [],
            monthlyIncome: getIngresos().data || {},
            config: getConfig()
        };
    } catch (error) {
        return { success: false, error: 'Error en getAllData: ' + error.toString() };
    }
}

function getExpenses() {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(SHEETS.GASTOS);
        
        if (!sheet) {
            return { success: false, error: `Hoja "${SHEETS.GASTOS}" no encontrada`, data: [] };
        }
        
        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return { success: true, data: [] };
        
        const numCols = Math.max(12, sheet.getLastColumn());
        const data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
        const expenses = [];
        
        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            
            if (!row[FORM_COLUMNS.FECHA] && !row[FORM_COLUMNS.MONTO]) continue;
            
            const fecha = formatDateForExport(row[FORM_COLUMNS.FECHA]);
            const monto = parseFloat(row[FORM_COLUMNS.MONTO]) || 0;
            const moneda = normalizeMoneda(row[FORM_COLUMNS.MONEDA]);
            const medioPago = normalizeMedioPago(row[FORM_COLUMNS.MEDIO_PAGO]);
            const categoria = normalizeCategoria(row[FORM_COLUMNS.CATEGORIA]);
            const descripcion = String(row[FORM_COLUMNS.DESCRIPCION] || '');
            
            // Columnas extendidas
            const cuotaActual = parseInt(row[8]) || 1;
            const totalCuotas = parseInt(row[9]) || parseInt(row[FORM_COLUMNS.CUOTAS]) || 1;
            const mesPagoGuardado = row[10] ? String(row[10]).trim() : '';
            const idGrupo = row[11] ? String(row[11]).trim() : null;  // NUEVO: leer idGrupo
            
            // Calcular mes de pago si no está guardado
            let mesPago = mesPagoGuardado;
            if (!mesPago) {
                mesPago = calcularMesPagoConCuota(fecha, medioPago, cuotaActual);
            }
            
            expenses.push({
                id: i + 1,
                rowIndex: i + 2,
                date: fecha,
                category: categoria,
                amount: monto,
                currency: moneda,
                payment: medioPago,
                description: descripcion,
                cuotaActual: cuotaActual,
                totalCuotas: totalCuotas,
                idGrupo: idGrupo,  // NUEVO: incluir en respuesta
                mesPago: mesPago
            });
        }
        
        return { success: true, data: expenses };
    } catch (error) {
        console.error('Error en getExpenses:', error);
        return { success: false, error: error.toString(), data: [] };
    }
}

/**
 * PATRIMONIO - Lectura desde hoja dedicada Patrimonio_App
 * Esto garantiza consistencia entre lectura y escritura
 */
function getPatrimonio() {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        
        // Primero intentar con la hoja dedicada de la app
        let sheet = ss.getSheetByName(SHEETS.PATRIMONIO_APP);
        
        if (sheet) {
            const data = sheet.getDataRange().getValues();
            const result = { bbva: 0, caja: 0, efectivo: 0 };
            
            for (let i = 1; i < data.length; i++) {
                const cuenta = String(data[i][0]).toLowerCase();
                const saldo = parseFloat(data[i][1]) || 0;
                
                if (cuenta.includes('bbva')) result.bbva = saldo;
                else if (cuenta.includes('caja')) result.caja = saldo;
                else if (cuenta.includes('efectivo')) result.efectivo = saldo;
            }
            
            return { success: true, data: result, source: 'Patrimonio_App' };
        }
        
        // Fallback a hoja Patrimonio legacy
        sheet = ss.getSheetByName(SHEETS.PATRIMONIO);
        if (sheet) {
            const data = sheet.getDataRange().getValues();
            const result = { bbva: 0, caja: 0, efectivo: 0 };
            
            for (let i = 1; i < data.length; i++) {
                const cuenta = String(data[i][0]).toLowerCase();
                if (cuenta.includes('bbva')) result.bbva = parseFloat(data[i][1]) || 0;
                else if (cuenta.includes('caja')) result.caja = parseFloat(data[i][1]) || 0;
                else if (cuenta.includes('efectivo')) result.efectivo = parseFloat(data[i][1]) || 0;
            }
            
            return { success: true, data: result, source: 'Patrimonio_Legacy' };
        }
        
        // Si no existe ninguna hoja, devolver valores default
        return { 
            success: true, 
            data: { bbva: 0, caja: 0, efectivo: 0 }, 
            source: 'Default',
            message: 'Ejecutá setupPatrimonio desde la app o ?action=setupPatrimonio'
        };
        
    } catch (error) {
        console.error('Error en getPatrimonio:', error);
        return { success: true, data: { bbva: 0, caja: 0, efectivo: 0 } };
    }
}

function getInversiones() {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(SHEETS.INVERSIONES);
        
        if (!sheet) return { success: true, data: [] };
        
        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return { success: true, data: [] };
        
        const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
        const inversiones = [];
        
        for (let i = 0; i < data.length; i++) {
            if (data[i][0]) {
                inversiones.push({
                    id: data[i][0],
                    nombre: data[i][1],
                    monto: parseFloat(data[i][2]) || 0,
                    tasa: parseFloat(data[i][3]) || 0,
                    frecuencia: data[i][4],
                    fechaCompra: formatDateForExport(data[i][5]),
                    vencimiento: formatDateForExport(data[i][6]),
                    origen: data[i][7] || ''
                });
            }
        }
        
        return { success: true, data: inversiones };
    } catch (error) {
        return { success: true, data: [] };
    }
}

function getMovimientos() {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(SHEETS.MOVIMIENTOS);
        
        if (!sheet) return { success: true, data: [] };
        
        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return { success: true, data: [] };
        
        const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
        const movimientos = [];
        
        for (let i = 0; i < data.length; i++) {
            if (data[i][0]) {
                movimientos.push({
                    id: data[i][0],
                    fecha: formatDateForExport(data[i][1]),
                    origen: data[i][2],
                    destino: data[i][3],
                    monto: parseFloat(data[i][4]) || 0,
                    nota: data[i][5] || '',
                    esAhorro: data[i][6] === 'TRUE' || data[i][6] === true
                });
            }
        }
        
        return { success: true, data: movimientos };
    } catch (error) {
        return { success: true, data: [] };
    }
}

function getIngresos() {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(SHEETS.INGRESOS);
        
        if (!sheet) return { success: true, data: {} };
        
        const data = sheet.getDataRange().getValues();
        const ingresos = {};
        
        for (let i = 1; i < data.length; i++) {
            if (data[i][0] && data[i][1]) {
                const key = `${data[i][1]}-${String(data[i][0]).padStart(2, '0')}`;
                ingresos[key] = parseFloat(data[i][2]) || 0;
            }
        }
        
        return { success: true, data: ingresos };
    } catch (error) {
        return { success: true, data: {} };
    }
}

function getConfig() {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(SHEETS.CONFIG);
        
        const config = { dolaresAhorro: 0 };
        
        if (sheet) {
            const data = sheet.getDataRange().getValues();
            for (let i = 1; i < data.length; i++) {
                if (data[i][0] === 'Dólares Ahorro') {
                    config.dolaresAhorro = parseFloat(data[i][1]) || 0;
                }
            }
        }
        
        return config;
    } catch (error) {
        return { dolaresAhorro: 0 };
    }
}

// ============================================================
// ESCRITURA - PATRIMONIO
// ============================================================

/**
 * Actualizar patrimonio - escribe en Patrimonio_App
 * Esta función crea la hoja si no existe
 */
function updatePatrimonio(patrimonio) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        let sheet = ss.getSheetByName(SHEETS.PATRIMONIO_APP);
        
        // Crear hoja si no existe
        if (!sheet) {
            sheet = ss.insertSheet(SHEETS.PATRIMONIO_APP);
            sheet.getRange('A1:C1').setValues([['Cuenta', 'Saldo USD', 'Última Actualización']]);
            sheet.getRange('A2:C4').setValues([
                ['BBVA', patrimonio.bbva || 0, new Date()],
                ['Caja Seguridad', patrimonio.caja || 0, new Date()],
                ['Efectivo', patrimonio.efectivo || 0, new Date()]
            ]);
            sheet.setColumnWidth(1, 150);
            sheet.setColumnWidth(2, 120);
            sheet.setColumnWidth(3, 180);
            return { success: true, created: true };
        }
        
        // Actualizar valores existentes
        const data = sheet.getDataRange().getValues();
        const now = new Date();
        
        for (let i = 1; i < data.length; i++) {
            const cuenta = String(data[i][0]).toLowerCase();
            const row = i + 1;
            
            if (cuenta.includes('bbva') && patrimonio.bbva !== undefined) {
                sheet.getRange(row, 2).setValue(patrimonio.bbva);
                sheet.getRange(row, 3).setValue(now);
            } else if (cuenta.includes('caja') && patrimonio.caja !== undefined) {
                sheet.getRange(row, 2).setValue(patrimonio.caja);
                sheet.getRange(row, 3).setValue(now);
            } else if (cuenta.includes('efectivo') && patrimonio.efectivo !== undefined) {
                sheet.getRange(row, 2).setValue(patrimonio.efectivo);
                sheet.getRange(row, 3).setValue(now);
            }
        }
        
        return { success: true };
    } catch (error) {
        console.error('Error en updatePatrimonio:', error);
        return { success: false, error: error.toString() };
    }
}

// ============================================================
// ESCRITURA - GASTOS
// ============================================================

function addExpense(expense) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(SHEETS.GASTOS);
        
        if (!sheet) return { success: false, error: 'Hoja de gastos no encontrada' };
        
        // Asegurar headers
        if (sheet.getLastColumn() < 12) setupGastosColumns();
        
        const row = [
            new Date(),                          // Timestamp
            expense.date,                        // Fecha
            expense.category,                    // Categoría
            expense.amount,                      // Monto
            expense.currency,                    // Moneda
            expense.payment,                     // Medio de pago
            1,                                   // Cuotas (form)
            expense.description,                 // Descripción
            1,                                   // Cuota Actual
            1,                                   // Total Cuotas
            expense.mesPago,                     // Mes Pago
            ''                                   // ID Grupo (vacío para gasto único)
        ];
        
        sheet.appendRow(row);
        const id = sheet.getLastRow() - 1;
        
        return { success: true, id: id };
    } catch (error) {
        return { success: false, error: error.toString() };
    }
}

/**
 * Agregar gasto con cuotas - genera idGrupo para agrupar todas las cuotas
 */
function addExpenseWithCuotas(expense, cuotas, mesesPago) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(SHEETS.GASTOS);
        
        if (!sheet) return { success: false, error: 'Hoja de gastos no encontrada' };
        
        // Asegurar headers
        if (sheet.getLastColumn() < 12) setupGastosColumns();
        
        const montoCuota = Math.round((expense.amount / cuotas) * 100) / 100;
        const groupId = 'GRP_' + Date.now();  // ID único para el grupo
        const ids = [];
        const rows = [];
        
        for (let i = 0; i < cuotas; i++) {
            const descripcionCuota = expense.description 
                ? `${expense.description} (cuota ${i + 1}/${cuotas})`
                : `${expense.category} (cuota ${i + 1}/${cuotas})`;
            
            rows.push([
                new Date(),                      // Timestamp
                expense.date,                    // Fecha original de compra
                expense.category,                // Categoría
                montoCuota,                      // Monto de la cuota
                expense.currency,                // Moneda
                expense.payment,                 // Medio de pago
                cuotas,                          // Cuotas (form original)
                descripcionCuota,                // Descripción con cuota
                i + 1,                           // Cuota Actual
                cuotas,                          // Total Cuotas
                mesesPago[i],                    // Mes Pago de esta cuota
                groupId                          // ID Grupo
            ]);
        }
        
        // Agregar todas las filas
        const startRow = sheet.getLastRow() + 1;
        sheet.getRange(startRow, 1, rows.length, 12).setValues(rows);
        
        // Generar IDs
        for (let i = 0; i < cuotas; i++) {
            ids.push(startRow + i - 1);
        }
        
        return { success: true, ids: ids, groupId: groupId };
    } catch (error) {
        console.error('Error en addExpenseWithCuotas:', error);
        return { success: false, error: error.toString() };
    }
}

function updateExpense(data) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(SHEETS.GASTOS);
        
        if (!sheet) return { success: false, error: 'Hoja no encontrada' };
        
        // Buscar por ID (rowIndex)
        const rowIndex = data.rowIndex || data.id + 1;
        
        if (rowIndex > 1 && rowIndex <= sheet.getLastRow()) {
            sheet.getRange(rowIndex, 2).setValue(data.date);
            sheet.getRange(rowIndex, 3).setValue(data.category);
            sheet.getRange(rowIndex, 4).setValue(data.amount);
            sheet.getRange(rowIndex, 5).setValue(data.currency);
            sheet.getRange(rowIndex, 6).setValue(data.payment);
            sheet.getRange(rowIndex, 8).setValue(data.description);
            if (data.mesPago) sheet.getRange(rowIndex, 11).setValue(data.mesPago);
            
            return { success: true };
        }
        
        return { success: false, error: 'Fila no encontrada' };
    } catch (error) {
        return { success: false, error: error.toString() };
    }
}

function deleteExpenseByRow(rowIndex) {
    try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.GASTOS);
        if (!sheet) return { success: false, error: 'Hoja no encontrada' };
        
        const row = parseInt(rowIndex);
        
        if (row > 1 && row <= sheet.getLastRow()) {
            sheet.deleteRow(row);
            return { success: true };
        }
        
        return { success: false, error: 'Fila no válida: ' + row };
    } catch (error) {
        return { success: false, error: error.toString() };
    }
}

// ============================================================
// ESCRITURA - INVERSIONES
// ============================================================
function addInversion(inv) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        let sheet = ss.getSheetByName(SHEETS.INVERSIONES);
        
        if (!sheet) {
            sheet = ss.insertSheet(SHEETS.INVERSIONES);
            sheet.getRange('A1:I1').setValues([['ID', 'Nombre', 'Monto USD', 'Tasa %', 'Frecuencia', 'Fecha Compra', 'Vencimiento', 'Origen', 'Notas']]);
        }
        
        const id = Date.now();
        sheet.appendRow([id, inv.nombre, inv.monto, inv.tasa, inv.frecuencia, inv.fechaCompra, inv.vencimiento, inv.origen || '', '']);
        
        return { success: true, id: id };
    } catch (error) {
        return { success: false, error: error.toString() };
    }
}

function updateInversion(inv) {
    try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.INVERSIONES);
        if (!sheet) return { success: false, error: 'Hoja no existe' };
        
        const data = sheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
            if (data[i][0] == inv.id) {
                const row = i + 1;
                sheet.getRange(row, 2).setValue(inv.nombre);
                sheet.getRange(row, 3).setValue(inv.monto);
                sheet.getRange(row, 4).setValue(inv.tasa);
                sheet.getRange(row, 5).setValue(inv.frecuencia);
                sheet.getRange(row, 6).setValue(inv.fechaCompra);
                sheet.getRange(row, 7).setValue(inv.vencimiento);
                sheet.getRange(row, 8).setValue(inv.origen || '');
                return { success: true };
            }
        }
        return { success: false, error: 'Inversión no encontrada' };
    } catch (error) {
        return { success: false, error: error.toString() };
    }
}

function deleteInversion(id) {
    try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.INVERSIONES);
        if (!sheet) return { success: false, error: 'Hoja no existe' };
        
        const data = sheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
            if (data[i][0] == id) {
                sheet.deleteRow(i + 1);
                return { success: true };
            }
        }
        return { success: false, error: 'Inversión no encontrada' };
    } catch (error) {
        return { success: false, error: error.toString() };
    }
}

// ============================================================
// ESCRITURA - MOVIMIENTOS
// ============================================================
function addMovimiento(mov) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        let sheet = ss.getSheetByName(SHEETS.MOVIMIENTOS);
        
        if (!sheet) {
            sheet = ss.insertSheet(SHEETS.MOVIMIENTOS);
            sheet.getRange('A1:G1').setValues([['ID', 'Fecha', 'Origen', 'Destino', 'Monto USD', 'Nota', 'Es Ahorro']]);
        }
        
        const id = Date.now();
        sheet.appendRow([id, mov.fecha, mov.origen, mov.destino, mov.monto, mov.nota || '', mov.esAhorro ? 'TRUE' : 'FALSE']);
        
        return { success: true, id: id };
    } catch (error) {
        return { success: false, error: error.toString() };
    }
}

// ============================================================
// ESCRITURA - INGRESOS
// ============================================================
function setMonthlyIncome(data) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        let sheet = ss.getSheetByName(SHEETS.INGRESOS);
        
        if (!sheet) {
            sheet = ss.insertSheet(SHEETS.INGRESOS);
            sheet.getRange('A1:C1').setValues([['Mes', 'Año', 'Monto']]);
        }
        
        const rows = sheet.getDataRange().getValues();
        
        for (let i = 1; i < rows.length; i++) {
            if (rows[i][0] == data.month && rows[i][1] == data.year) {
                sheet.getRange(i + 1, 3).setValue(data.amount);
                return { success: true };
            }
        }
        
        sheet.appendRow([data.month, data.year, data.amount]);
        return { success: true };
    } catch (error) {
        return { success: false, error: error.toString() };
    }
}

// ============================================================
// ESCRITURA - CONFIG
// ============================================================
function setDolaresAhorro(amount) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        let sheet = ss.getSheetByName(SHEETS.CONFIG);
        
        if (!sheet) {
            sheet = ss.insertSheet(SHEETS.CONFIG);
            sheet.getRange('A1:B1').setValues([['Parámetro', 'Valor']]);
            sheet.appendRow(['Dólares Ahorro', amount]);
            return { success: true };
        }
        
        const data = sheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
            if (data[i][0] === 'Dólares Ahorro') {
                sheet.getRange(i + 1, 2).setValue(amount);
                return { success: true };
            }
        }
        
        sheet.appendRow(['Dólares Ahorro', amount]);
        return { success: true };
    } catch (error) {
        return { success: false, error: error.toString() };
    }
}

// ============================================================
// UTILIDADES DE FORMATO
// ============================================================
function formatDateForExport(value) {
    if (!value) return '';
    if (value instanceof Date) {
        return Utilities.formatDate(value, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd');
    }
    const str = String(value);
    if (str.includes('T')) return str.split('T')[0];
    if (str.includes(' ')) return str.split(' ')[0];
    return str;
}

function normalizeMoneda(value) {
    const str = String(value || 'ARS').toUpperCase();
    if (str.includes('USD') || str.includes('DOLAR') || str.includes('DÓLAR')) return 'USD';
    return 'ARS';
}

function normalizeMedioPago(value) {
    const str = String(value || '').toLowerCase();
    if (str.includes('visa')) return 'VISA (8043)';
    if (str.includes('master')) return 'MASTER (9714)';
    if (str.includes('débito') || str.includes('debito') || str.includes('transfere')) return 'Débito/Transferencia';
    if (str.includes('efectivo')) return 'Efectivo';
    return value || 'Efectivo';
}

function normalizeCategoria(value) {
    const str = String(value || 'Otros');
    return str.charAt(0).toUpperCase() + str.slice(1).toLowerCase();
}

function calcularMesPago(fechaStr, medioPago) {
    if (!fechaStr) return '';
    
    try {
        const fecha = new Date(fechaStr + 'T12:00:00');
        const year = fecha.getFullYear();
        const month = fecha.getMonth();
        
        const isTarjeta = medioPago && (medioPago.includes('VISA') || medioPago.includes('MASTER'));
        
        if (isTarjeta) {
            // Calcular último jueves del mes (cierre de tarjeta)
            const lastDay = new Date(year, month + 1, 0);
            const dayOfWeek = lastDay.getDay();
            const diff = (dayOfWeek >= 4) ? (dayOfWeek - 4) : (dayOfWeek + 3);
            const lastThursday = new Date(year, month + 1, 0 - diff);
            
            if (fecha <= lastThursday) {
                const vtoMonth = month + 1;
                if (vtoMonth > 11) return `${year + 1}-01`;
                return `${year}-${String(vtoMonth + 1).padStart(2, '0')}`;
            } else {
                const vtoMonth = month + 2;
                if (vtoMonth > 11) return `${year + 1}-${String(vtoMonth - 11).padStart(2, '0')}`;
                return `${year}-${String(vtoMonth + 1).padStart(2, '0')}`;
            }
        } else {
            return `${year}-${String(month + 1).padStart(2, '0')}`;
        }
    } catch (e) {
        return '';
    }
}

function calcularMesPagoConCuota(fechaStr, medioPago, cuotaActual) {
    if (!fechaStr) return '';
    
    try {
        const fecha = new Date(fechaStr + 'T12:00:00');
        const year = fecha.getFullYear();
        const month = fecha.getMonth();
        
        const isTarjeta = medioPago && (medioPago.includes('VISA') || medioPago.includes('MASTER') || medioPago.toLowerCase().includes('tarjeta'));
        
        if (isTarjeta) {
            const lastDay = new Date(year, month + 1, 0);
            const dayOfWeek = lastDay.getDay();
            const diff = (dayOfWeek >= 4) ? (dayOfWeek - 4) : (dayOfWeek + 3);
            const lastThursday = new Date(year, month + 1, 0 - diff);
            
            let baseMonth, baseYear;
            if (fecha <= lastThursday) {
                baseMonth = month + 1;
                baseYear = year;
            } else {
                baseMonth = month + 2;
                baseYear = year;
            }
            
            const cuotaOffset = (cuotaActual || 1) - 1;
            let finalMonth = baseMonth + cuotaOffset;
            let finalYear = baseYear;
            
            while (finalMonth > 11) {
                finalMonth -= 12;
                finalYear++;
            }
            
            return `${finalYear}-${String(finalMonth + 1).padStart(2, '0')}`;
        } else {
            return `${year}-${String(month + 1).padStart(2, '0')}`;
        }
    } catch (e) {
        return '';
    }
}

function calcularMesesPagoCuotas(fechaCompra, cuotas, medioPago) {
    const meses = [];
    const fecha = new Date(fechaCompra);
    const year = fecha.getFullYear();
    const month = fecha.getMonth();
    
    const lastDay = new Date(year, month + 1, 0);
    const dayOfWeek = lastDay.getDay();
    const diff = (dayOfWeek >= 4) ? (dayOfWeek - 4) : (dayOfWeek + 3);
    const lastThursday = new Date(year, month + 1, 0 - diff);
    
    let baseMonth, baseYear;
    if (fecha <= lastThursday) {
        baseMonth = month + 1;
        baseYear = year;
    } else {
        baseMonth = month + 2;
        baseYear = year;
    }
    
    for (let i = 0; i < cuotas; i++) {
        let m = baseMonth + i;
        let y = baseYear;
        while (m > 11) {
            m -= 12;
            y++;
        }
        meses.push(`${y}-${String(m + 1).padStart(2, '0')}`);
    }
    
    return meses;
}

// ============================================================
// TRIGGER PARA EXPANDIR CUOTAS (Google Forms)
// ============================================================
function onFormSubmit(e) {
    try {
        const sheet = e.range.getSheet();
        
        if (sheet.getName() !== SHEETS.GASTOS) return;
        
        const row = e.range.getRow();
        const values = sheet.getRange(row, 1, 1, 8).getValues()[0];
        
        const fecha = values[FORM_COLUMNS.FECHA];
        const cuotas = parseInt(values[FORM_COLUMNS.CUOTAS]) || 1;
        
        // Asegurar que existan las columnas extendidas
        if (sheet.getLastColumn() < 12) {
            sheet.getRange(1, 9).setValue('Cuota');
            sheet.getRange(1, 10).setValue('Total Cuotas');
            sheet.getRange(1, 11).setValue('Mes Pago');
            sheet.getRange(1, 12).setValue('ID Grupo');
        }
        
        if (cuotas > 1) {
            const monto = parseFloat(values[FORM_COLUMNS.MONTO]) || 0;
            const montoCuota = Math.round((monto / cuotas) * 100) / 100;
            const medioPago = values[FORM_COLUMNS.MEDIO_PAGO];
            const descripcion = values[FORM_COLUMNS.DESCRIPCION];
            const groupId = 'GRP_' + Date.now();  // ID único para el grupo
            
            const mesesPago = calcularMesesPagoCuotas(fecha, cuotas, medioPago);
            
            // Modificar fila original (cuota 1)
            sheet.getRange(row, FORM_COLUMNS.MONTO + 1).setValue(montoCuota);
            sheet.getRange(row, 9).setValue(1);
            sheet.getRange(row, 10).setValue(cuotas);
            sheet.getRange(row, 11).setValue(mesesPago[0]);
            sheet.getRange(row, 12).setValue(groupId);
            sheet.getRange(row, FORM_COLUMNS.DESCRIPCION + 1).setValue(descripcion + ' (cuota 1/' + cuotas + ')');
            
            // Agregar filas para cuotas 2 en adelante
            const newRows = [];
            for (let i = 1; i < cuotas; i++) {
                newRows.push([
                    values[FORM_COLUMNS.TIMESTAMP],
                    fecha,
                    values[FORM_COLUMNS.CATEGORIA],
                    montoCuota,
                    values[FORM_COLUMNS.MONEDA],
                    medioPago,
                    cuotas,
                    descripcion + ' (cuota ' + (i + 1) + '/' + cuotas + ')',
                    i + 1,
                    cuotas,
                    mesesPago[i],
                    groupId
                ]);
            }
            
            if (newRows.length > 0) {
                sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 12).setValues(newRows);
            }
        } else {
            // Una sola cuota
            const medioPago = values[FORM_COLUMNS.MEDIO_PAGO];
            const mesPago = calcularMesPago(formatDateForExport(fecha), normalizeMedioPago(medioPago));
            sheet.getRange(row, 9).setValue(1);
            sheet.getRange(row, 10).setValue(1);
            sheet.getRange(row, 11).setValue(mesPago);
            sheet.getRange(row, 12).setValue('');
        }
    } catch (error) {
        console.error('Error en onFormSubmit:', error);
    }
}

// ============================================================
// DEBUG
// ============================================================
function debugSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets().map(s => s.getName());
    
    const result = {
        success: true,
        version: '5.3',
        spreadsheetName: ss.getName(),
        sheets: sheets,
        hojaGastosConfigurada: SHEETS.GASTOS,
        hojaPatrimonioConfigurada: SHEETS.PATRIMONIO_APP
    };
    
    // Verificar hoja de gastos
    const gastosSheet = ss.getSheetByName(SHEETS.GASTOS);
    if (gastosSheet) {
        result.hojaGastosEncontrada = true;
        result.gastosFilas = gastosSheet.getLastRow();
        result.gastosColumnas = gastosSheet.getLastColumn();
        if (gastosSheet.getLastRow() >= 1) {
            result.gastosHeaders = gastosSheet.getRange(1, 1, 1, Math.min(12, gastosSheet.getLastColumn())).getValues()[0];
        }
    } else {
        result.hojaGastosEncontrada = false;
    }
    
    // Verificar hoja de patrimonio
    const patrimonioSheet = ss.getSheetByName(SHEETS.PATRIMONIO_APP);
    if (patrimonioSheet) {
        result.hojaPatrimonioEncontrada = true;
        result.patrimonioData = patrimonioSheet.getDataRange().getValues();
    } else {
        result.hojaPatrimonioEncontrada = false;
        result.sugerencia = 'Ejecutar setupPatrimonio para crear la hoja';
    }
    
    return result;
}
