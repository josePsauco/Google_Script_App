function onOpen(e){
    DocumentApp.getUi().createMenu('Menu-Operaciones').addItem('Cargar Tabla excel', 'Capturar').addItem('Enviar mensaje', 'EnviarMensaje').addItem('Enviar PDF', 'EnviarPDF').addItem('Ver Cuerpo del mensaje', 'VerMensaje').
    addItem('Conectar hoja de calculo', 'ConectarDocumento').addToUi();
  }
  
  function rellenarCeldas(sheet) {
    var hojaSheet = sheet.getActiveSheet();
    var ultimaCD = Number(sheet.getSheets()[0].getLastColumn());
    var ultimaFD = Number(sheet.getSheets()[0].getLastRow());
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    var table = body.appendTable();  
    var row1 = table.appendTableRow();
    var col1row1 = row1.appendTableCell("Código");
    var col2row1 = row1.appendTableCell("Nombres");
    var col3row1 = row1.appendTableCell("Apellidos");
    var col4row1 = row1.appendTableCell("Calificaciones");
    var col5row1 = row1.appendTableCell("Observación");
    var col6row1 = row1.appendTableCell("Email");
    col1row1.setBackgroundColor("#D8D8D8");
    col2row1.setBackgroundColor("#D8D8D8");
    col3row1.setBackgroundColor("#D8D8D8");
    col4row1.setBackgroundColor("#D8D8D8");
    col5row1.setBackgroundColor("#D8D8D8");
    col6row1.setBackgroundColor("#D8D8D8");
    
    for(var row =2 ; row <= ultimaFD; row++){
      var row2 = table.appendTableRow();
      for(var col =1 ; col <= ultimaCD; col++){
          var celdaImportada = hojaSheet.getRange(row, col);
          row2.appendTableCell(celdaImportada.getValue());
      }
    }
  }
  
  
  function Capturar() {
    
    var html = HtmlService.createHtmlOutputFromFile('Archivo.html');
    html.setHeight(60);
    html.setWidth(500);
    DocumentApp.getUi().showModalDialog(html,'Adjunte documento excel');
  }
  
  
  function CargarArchivo(obj) {
    if(obj!=null){
      var blob = Utilities.newBlob(Utilities.base64Decode(obj.data), obj.mimeType, obj.fileName);
      var id  = ConvertirGuardar(blob);
      if(id!=null){Logger.log(id);
                   var documento = SpreadsheetApp.openById(id);
                   rellenarCeldas(documento);
                   return;
                  }
      return
    }
    Error();
      
  }
  
  
  function ConvertirGuardar(blob) {
  
    try {
      var resource = {
        title: blob.getName(),
        mimeType: MimeType.GOOGLE_SHEETS
      };
      
      var id = Drive.Files.insert(resource, blob);
      return id.id;
    } catch (f) {
      return null;
    }
  
  }
  
  
  function ArmarMatrix() {
    
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    var tabla = body.getTables()[0];
    if(tabla!=null){
      var colum = 6;
      var filas = tabla.getNumRows();
      var matrix = [];
      for(var i=0; i<filas; i++) {
        matrix[i] = new Array(colum);
        for(var j=0; j<colum; j++) {
          matrix[i][j] = tabla.getCell(i, j).getText();
        }
      }
      return matrix}
    return null;
  }
  
  function EnviarMensaje() {
    
    var matrix = ArmarMatrix();
    if(matrix!=null){
      for(var i=1; i<matrix.length; i++) {
        GmailApp.sendEmail(matrix[i][5], "Combinacion de correspondecia",
                           "Estimado (a) estudiante: "+matrix[i][1]+" "+matrix[i][2]+" \nNos permite informarle que su reporte en el curso Ing.\nCalificacion: "+matrix[i][3]+
                           ".\nObservaciones: "+matrix[i][4]+".\nPuesto en el grupo: "+matrix[i][0]+".\nFelicitaciones y exitos en su vida academica y profesional.");
      }
      return;}
    Error();
  }
  
  function EnviarPDF(){
    var doc = DocumentApp.getActiveDocument();
    var fileDoc = DriveApp.getFileById(doc.getId());
    var pdfDoc = DriveApp.createFile(fileDoc.getAs('application/pdf'));
    pdfDoc.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    var recipient = Session.getActiveUser().getEmail();
    GmailApp.sendEmail(recipient, 'Listado de estudiante', 'URL PDF ' + pdfDoc.getUrl());
  }
  
  function VerMensaje(){
    var html = HtmlService.createHtmlOutputFromFile('Cuerpo.html');
    html.setHeight(300);
    html.setWidth(500);
    DocumentApp.getUi().showModalDialog(html,'Cuerpo del Email');
  }
  
  function Error(){
    var html = HtmlService.createHtmlOutputFromFile('Error.html');
    html.setHeight(30);
    html.setWidth(500);
    DocumentApp.getUi().showModalDialog(html,'Error');
  }
  
  function ConectarDocumento(){
    
    var html = HtmlService.createHtmlOutputFromFile('Interfaz.html');
    html.setHeight(30);
    html.setWidth(500);
    DocumentApp.getUi().showModalDialog(html,'Ingrese la url de la hoja de google a conectar');
  }

  function BuscarHojaGoogle(obj){
    var documento = SpreadsheetApp.openByUrl(obj.file);
    if(documento!=null){
      rellenarCeldas(documento);
      return;}
    Error();
  }
  
