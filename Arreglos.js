
function leerArreglos() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = libro.getActiveSheet();
  var arregloDatos = hoja.getDataRange().getValues();
  console.log(arregloDatos)
  
  for(var fila=1;fila <= arregloDatos.length-1;fila++)
  {
    var tipo=arregloDatos[fila][2]
    var sS = arregloDatos[fila][1]     
    Logger.log(sS)
    Logger.log(tipo)
  }
  
}
function leerArreglosForEach() 
{
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = libro.getActiveSheet();
  var arregloDatos = hoja.getDataRange().getValues();
  console.log(arregloDatos)
  arregloDatos.forEach(fila=>{
    var tipo=[fila][2]
    var sS = [fila][1]     
    Logger.log(sS)
    Logger.log(tipo)

  })

  }
  //Tomar este para hacer filtrar la BD y copiar a la hoja Resultado
  function macro_1() {
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
 otra forma
  var areaC ="Resultad"
  var targetSheet = libro.insertSheet(areaC);
 targetSheet.getRange(2,1,data.length,data[0].length).setValues(data);
*/

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
 DriveApp.createFile(xlsxBlob.setName(`${name}.xlsx`));

  // 6. Delete the temporate Spreadsheet.
  DriveApp.getFileById(id).setTrashed(true);
 
}

//Operadores Logicos || (o)  && (y)
/*var filterlogic=function(item){
  if(item[3]==="Facturación")
  {
    return true;
  } else{
    return false;
  }

}*/
