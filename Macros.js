function eliminar_data_BD()
{
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = libro.getSheetByName("BD");
  hoja.getRange(1,1,hoja.getLastRow(),hoja.getLastColumn()).clearContent();
}  
  
  
//Tomar este para hacer filtrar la BD y copiar a la hoja Resultado
  function obtener_Excel() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = libro.getSheetByName("BD");
  var originalData = hoja.getRange(2,1,hoja.getLastRow()-1,39).getValues();
  //var data =originalData.filter(filterlogic); Una forma de hacerlo con una variable global
  var hojaDestinoR = libro.getSheetByName("Resultado");
  var data =originalData.filter(function(item){return item[3]==="Facturación" && item[37]>51900000000 && item[37]<51999999999 });// con una sola linea
 Logger.log(data)

  hojaDestinoR.getRange(2,1,hojaDestinoR.getLastRow(),hojaDestinoR.getLastColumn()).clearContent();
  hojaDestinoR.getRange(2,1,data.length,data[0].length).setValues(data);
/*
Guardamos la hoja en formato xlsx en drive con el nombre  de la fecha y hora exacta que se ejecuta
*/

 const exportSheetName = 'Resultado';  // Please set the target sheet name.
  
  // 1. Copy the active Spreadsheet as a tempora Spreadsheet.
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet().copy('tmp');

  // 2. Convert the formulas to the texts.
  const targetRange = spreadsheet.getSheetByName(exportSheetName).getDataRange();
  targetRange.copyTo(targetRange, {contentsOnly:true});

  // 3. Delete the sheets except for a sheet you want to export.
  spreadsheet.getSheets().forEach(sheet => {
    if (exportSheetName != sheet.getName()) spreadsheet.deleteSheet(sheet)
  });

  // 4. Retrieve the blob from the export URL.
  const id = spreadsheet.getId();
  const xlsxBlob = UrlFetchApp.fetch(`https://docs.google.com/spreadsheets/export?id=${id}&exportFormat=xlsx`, {headers: {authorization: `Bearer ${ScriptApp.getOAuthToken()}`}}).getBlob();
  
  // 5. Crete the blob as a file.
 // DriveApp.createFile(xlsxBlob.setName(`${exportSheetName}.xlsx`));
 
 //Restando 5 horas a la hora amaricana por default
var five_Hours = 1000 * 60 * 60 * 5;
var now = new Date();
var day_Before_Five_Hours = new Date(now.getTime() - five_Hours);

var formattedDate=Utilities.formatDate(day_Before_Five_Hours,'ETC/GMT',"yyyy-MM-dd' 'HH:mm:ss")
var name=SpreadsheetApp.getActiveSpreadsheet().getName()+"_"+formattedDate;
 DriveApp.getFolderById('1n2Vy2qc-bY5P1YKtcAVIqqabC5uS_MZF').createFile(xlsxBlob.setName(`${name}.xlsx`));

  // 6. Delete the temporate Spreadsheet.
  DriveApp.getFileById(id).setTrashed(true);
 
}
function obtener_Csv(){
var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = libro.getSheetByName("Resultado");
  var originalData = hoja.getRange(2, 1, hoja.getLastRow() - 1, 39).getValues();
  var csvData = [];
  //Agregar fila de encabezado
  var headerRow = [
    'MSISDN',
    'CLASIFICACION',
    'TIPO',
    'FECHA_HORA',
    'CANAL 1-1',
    'CANAL 1-2',
    'CANAL 2',
    'PRIORIDAD',
    'MENSAJE',
    'ACCION'

  ];
  csvData.push(headerRow.join('|'));

  originalData.forEach(function (item) {
    var msisdn = item[1];
    var mensaje = "Hola, hemos procedido a ejecutar la solucion anticipada de reclamo a tu favor Nro "+ `${item[37]}`+" , si tienes dudas, escribenos a nuestro Whatsapp 981002000";
    var fechaHora = new Date();
    fechaHora.setHours(fechaHora.getHours() + 1); // Agregar una hora a la fecha actual
    var fila = [
      msisdn,
      'ACUSE',
      'Entel',
      formatDate(fechaHora),
      'SMS',
      '',
      '',
      '1',
      mensaje,
      ''
    ];
    csvData.push(fila.join('|'));
  });

  var csvString = csvData.join('\n');
  var nombreArchivo='SAR_'+formatDateForFileName(new Date());

  // Exportar como archivo CSV
  var folder = DriveApp.getFolderById('1n2Vy2qc-bY5P1YKtcAVIqqabC5uS_MZF');
  var csvFile = folder.createFile(nombreArchivo, csvString , MimeType.CSV);
}


// Función auxiliar para formatear la fecha y hora
function formatDate(date) {
  var year = date.getFullYear();
  var month = padZero(date.getMonth() + 1);
  var day = padZero(date.getDate());
  var hours = padZero(date.getHours());
  var minutes = padZero(date.getMinutes());
  var seconds = padZero(date.getSeconds());
  return year + '-' + month + '-' + day + ' ' + hours + ':' + minutes + ':' + seconds;
}
function formatDateForFileName(date) {
  var year = date.getFullYear();
  var month = padZero(date.getMonth() + 1);
  var day = padZero(date.getDate());
  var hours = padZero(date.getHours());
  var minutes = padZero(date.getMinutes());
  var seconds = padZero(date.getSeconds());
  return year  + month + day + hours + minutes + seconds;
}

// Función auxiliar para agregar ceros a la izquierda si es necesario
function padZero(value) {
  return value.toString().padStart(2, '0');
}