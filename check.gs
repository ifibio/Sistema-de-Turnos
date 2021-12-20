function obtener_turnos() {
  var origen = SpreadsheetApp.openById("ID_PLANILLA").getSheetByName("Consulta de asistencias");
  var sector_id = origen.getRange("D3").getDisplayValues();
  var desde = origen.getRange("B3").getDisplayValues();
  var hasta = origen.getRange("B6").getDisplayValues();

  var cal = CalendarApp.getCalendarById(sector_id);

  var events = cal.getEvents(new Date(desde), new Date(hasta));


  var sheet = SpreadsheetApp.openById("ID_PLANILLA").getSheetByName("Consulta de asistencias");
  // Uncomment this next line if you want to always clear the spreadsheet content before running - Note people could have added extra columns on the data though that would be lost

  var lastRow = sheet.getLastRow();

  // Rows start at "1" - this deletes the first two rows
  sheet.deleteRows(7, lastRow);

  // Create a header record on the current spreadsheet in cells A1:N1 - Match the number of entries in the "header=" to the last parameter
  // of the getRange entry below
  var header = [["Titulo", "Descripción", "Sector", "Inicio", "Final", "Duración (Horas)", "Creado", "Estado", "Creado por", "Se repite"]];
  var range = sheet.getRange(8, 1, 1, 10);
  range.setValues(header); 
  range.setFontWeight("bold");


  // Loop eventos en el rango de fechas
  for (var i = 0; i < events.length; i++) {
    var row = i + 9;

    var details = [[events[i].getTitle(), events[i].getDescription(), events[i].getLocation(), events[i].getStartTime(), events[i].getEndTime(), '', events[i].getDateCreated(), events[i].getMyStatus(), events[i].getCreators(), events[i].isRecurringEvent()]];

    var range = sheet.getRange(row, 1, 1, 10);
    range.setValues(details);

    var cell_duracion = sheet.getRange(row, 6);

    // d/m/yy at h:mm

    // row.setNumberFormat('d/m/yy at h:mm');
    var col_inicio = sheet.getRange(row, 4);
    var col_final = sheet.getRange(row, 5);
    col_inicio.setNumberFormat('d/m/yy at h:mm');
    col_final.setNumberFormat('d/m/yy at h:mm');

    var inicio = events[i].getStartTime();
    var final = events[i].getEndTime();

    // tiempo en milisegundos

    var t_inicio = inicio.getTime();
    var t_final = final.getTime();

    var duracion = (t_final - t_inicio) / (3600 * 1000);

    cell_duracion.setValue(duracion);
    cell_duracion.setNumberFormat('00.00');

  }

}



