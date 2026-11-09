function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Soporte Informático - Colegio San Gabriel')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Obtener o crear la hoja de Tareas
function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Tareas');
  if (!sheet) {
    sheet = ss.insertSheet('Tareas');
    sheet.appendRow(['ID', 'Título', 'Descripción', 'Lugar', 'Responsable', 'Prioridad', 'Fecha Creación', 'Fecha Ejecución', 'Hora Ejecución', 'Estado', 'EvidenciaBase64']);
  }
  return sheet;
}

// Cargar todas las tareas
function getTasksFromSheet() {
  try {
    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    var tasks = [];
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        tasks.push({
          id: data[i][0].toString(),
          title: data[i][1] ? data[i][1].toString() : '',
          description: data[i][2] ? data[i][2].toString() : '',
          location: data[i][3] ? data[i][3].toString() : '',
          assignee: data[i][4] ? data[i][4].toString() : '',
          priority: data[i][5] ? data[i][5].toString() : 'Media',
          createdDate: data[i][6] ? formatDate(data[i][6]) : '',
          execDate: data[i][7] ? formatDate(data[i][7]) : '',
          execTime: data[i][8] ? data[i][8].toString() : '09:00',
          status: data[i][9] ? data[i][9].toString() : 'Pendiente',
          evidence: data[i][10] ? data[i][10].toString() : ''
        });
      }
    }
    return tasks;
  } catch(e) {
    Logger.log('Error en getTasksFromSheet: ' + e.message);
    return [];
  }
}

// Guardar o actualizar tarea
function saveTaskToSheet(task) {
  try {
    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === task.id.toString()) {
        rowIndex = i + 1;
        break;
      }
    }

    var rowPayload = [
      task.id || Date.now().toString(), 
      task.title || '', 
      task.description || '', 
      task.location || '', 
      task.assignee || '', 
      task.priority || 'Media', 
      task.createdDate || formatDate(new Date()), 
      task.execDate || formatDate(new Date()), 
      task.execTime || '09:00', 
      task.status || 'Pendiente', 
      task.evidence || ''
    ];

    if (rowIndex > -1) {
      sheet.getRange(rowIndex, 1, 1, 11).setValues([rowPayload]);
    } else {
      sheet.appendRow(rowPayload);
    }
    return true;
  } catch(e) {
    Logger.log('Error en saveTaskToSheet: ' + e.message);
    throw new Error(e.message);
  }
}

// Eliminar tarea
function deleteTaskFromSheet(taskId) {
  try {
    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === taskId.toString()) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    return true;
  } catch(e) {
    return false;
  }
}

// Formatear fechas
function formatDate(date) {
  if (!date) return '';
  if (typeof date === 'string') return date.substring(0, 10);
  try {
    var tz = Session.getScriptTimeZone();
    return Utilities.formatDate(new Date(date), tz, 'yyyy-MM-dd');
  } catch (e) {
    return '';
  }
}

// Obtener o crear la hoja de Solicitudes
function getRequestsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('SolicitudesDocentes');
  if (!sheet) {
    sheet = ss.insertSheet('SolicitudesDocentes');
    sheet.appendRow(['ID', 'Docente', 'Ubicacion', 'Dispositivo', 'Descripcion', 'Prioridad', 'Fecha', 'Estado']);
  }
  return sheet;
}

function getTeacherRequests() {
  try {
    var sheet = getRequestsSheet();
    var data = sheet.getDataRange().getValues();
    var requests = [];
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        requests.push({
          id: data[i][0].toString(),
          teacher: data[i][1] ? data[i][1].toString() : '',
          location: data[i][2] ? data[i][2].toString() : '',
          device: data[i][3] ? data[i][3].toString() : '',
          description: data[i][4] ? data[i][4].toString() : '',
          priority: data[i][5] ? data[i][5].toString() : 'Media',
          date: data[i][6] ? formatDate(data[i][6]) : '',
          status: data[i][7] ? data[i][7].toString() : 'Pendiente'
        });
      }
    }
    return requests;
  } catch(e) {
    Logger.log('Error en getTeacherRequests: ' + e.message);
    return [];
  }
}

function saveTeacherRequest(req) {
  try {
    var sheet = getRequestsSheet();
    sheet.appendRow([
      req.id || Date.now().toString(),
      req.teacher || '',
      req.location || '',
      req.device || '',
      req.description || '',
      req.priority || 'Media',
      req.date || formatDate(new Date()),
      req.status || 'Pendiente'
    ]);
    return true;
  } catch(e) {
    Logger.log('Error en saveTeacherRequest: ' + e.message);
    throw new Error(e.message);
  }
}

function updateRequestStatus(reqId, newStatus) {
  try {
    var sheet = getRequestsSheet();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === reqId.toString()) {
        sheet.getRange(i + 1, 8).setValue(newStatus);
        break;
      }
    }
    return true;
  } catch(e) {
    return false;
  }
}function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Soporte Informático - Colegio San Gabriel')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Obtener o crear la hoja de Tareas
function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Tareas');
  if (!sheet) {
    sheet = ss.insertSheet('Tareas');
    sheet.appendRow(['ID', 'Título', 'Descripción', 'Lugar', 'Responsable', 'Prioridad', 'Fecha Creación', 'Fecha Ejecución', 'Hora Ejecución', 'Estado', 'EvidenciaBase64']);
  }
  return sheet;
}

// Cargar todas las tareas
function getTasksFromSheet() {
  try {
    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    var tasks = [];
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        tasks.push({
          id: data[i][0].toString(),
          title: data[i][1] ? data[i][1].toString() : '',
          description: data[i][2] ? data[i][2].toString() : '',
          location: data[i][3] ? data[i][3].toString() : '',
          assignee: data[i][4] ? data[i][4].toString() : '',
          priority: data[i][5] ? data[i][5].toString() : 'Media',
          createdDate: data[i][6] ? formatDate(data[i][6]) : '',
          execDate: data[i][7] ? formatDate(data[i][7]) : '',
          execTime: data[i][8] ? data[i][8].toString() : '09:00',
          status: data[i][9] ? data[i][9].toString() : 'Pendiente',
          evidence: data[i][10] ? data[i][10].toString() : ''
        });
      }
    }
    return tasks;
  } catch(e) {
    Logger.log('Error en getTasksFromSheet: ' + e.message);
    return [];
  }
}

// Guardar o actualizar tarea
function saveTaskToSheet(task) {
  try {
    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === task.id.toString()) {
        rowIndex = i + 1;
        break;
      }
    }

    var rowPayload = [
      task.id || Date.now().toString(), 
      task.title || '', 
      task.description || '', 
      task.location || '', 
      task.assignee || '', 
      task.priority || 'Media', 
      task.createdDate || formatDate(new Date()), 
      task.execDate || formatDate(new Date()), 
      task.execTime || '09:00', 
      task.status || 'Pendiente', 
      task.evidence || ''
    ];

    if (rowIndex > -1) {
      sheet.getRange(rowIndex, 1, 1, 11).setValues([rowPayload]);
    } else {
      sheet.appendRow(rowPayload);
    }
    return true;
  } catch(e) {
    Logger.log('Error en saveTaskToSheet: ' + e.message);
    throw new Error(e.message);
  }
}

// Eliminar tarea
function deleteTaskFromSheet(taskId) {
  try {
    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === taskId.toString()) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    return true;
  } catch(e) {
    return false;
  }
}

// Formatear fechas
function formatDate(date) {
  if (!date) return '';
  if (typeof date === 'string') return date.substring(0, 10);
  try {
    var tz = Session.getScriptTimeZone();
    return Utilities.formatDate(new Date(date), tz, 'yyyy-MM-dd');
  } catch (e) {
    return '';
  }
}

// Obtener o crear la hoja de Solicitudes
function getRequestsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('SolicitudesDocentes');
  if (!sheet) {
    sheet = ss.insertSheet('SolicitudesDocentes');
    sheet.appendRow(['ID', 'Docente', 'Ubicacion', 'Dispositivo', 'Descripcion', 'Prioridad', 'Fecha', 'Estado']);
  }
  return sheet;
}

function getTeacherRequests() {
  try {
    var sheet = getRequestsSheet();
    var data = sheet.getDataRange().getValues();
    var requests = [];
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        requests.push({
          id: data[i][0].toString(),
          teacher: data[i][1] ? data[i][1].toString() : '',
          location: data[i][2] ? data[i][2].toString() : '',
          device: data[i][3] ? data[i][3].toString() : '',
          description: data[i][4] ? data[i][4].toString() : '',
          priority: data[i][5] ? data[i][5].toString() : 'Media',
          date: data[i][6] ? formatDate(data[i][6]) : '',
          status: data[i][7] ? data[i][7].toString() : 'Pendiente'
        });
      }
    }
    return requests;
  } catch(e) {
    Logger.log('Error en getTeacherRequests: ' + e.message);
    return [];
  }
}

function saveTeacherRequest(req) {
  try {
    var sheet = getRequestsSheet();
    sheet.appendRow([
      req.id || Date.now().toString(),
      req.teacher || '',
      req.location || '',
      req.device || '',
      req.description || '',
      req.priority || 'Media',
      req.date || formatDate(new Date()),
      req.status || 'Pendiente'
    ]);
    return true;
  } catch(e) {
    Logger.log('Error en saveTeacherRequest: ' + e.message);
    throw new Error(e.message);
  }
}

function updateRequestStatus(reqId, newStatus) {
  try {
    var sheet = getRequestsSheet();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === reqId.toString()) {
        sheet.getRange(i + 1, 8).setValue(newStatus);
        break;
      }
    }
    return true;
  } catch(e) {
    return false;
  }
}
