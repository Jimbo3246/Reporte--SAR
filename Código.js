
function copiarfirstBD_FILTRADATODescargar() 
{ const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hojaOrigen = libro.getActiveSheet();
  const hojaDestino = libro.getSheetByName("Descargar");
  const rangoOrigen = hojaOrigen.getRange(1,1,hojaOrigen.getLastRow(),hojaOrigen.getLastColumn()).getValues()
  const rangoDestino = hojaDestino.getRange(1,1,hojaOrigen.getLastRow(),hojaOrigen.getLastColumn()).setValues(rangoOrigen)
}
/*
Esta funcion sirve para guardar el archivo de una hoja en especifico  en el drive 
como  formato xlsx con el nombre en la hora y fecha
*/
function macro_2() {
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
 DriveApp.createFile(xlsxBlob.setName(`${name}.xlsx`));

  // 6. Delete the temporate Spreadsheet.
  DriveApp.getFileById(id).setTrashed(true);
}
