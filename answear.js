function calculateAndWriteResults() {
    // Access the active spreadsheet
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName('engenharia_de_software'); // Replace 'engenharia_de_software' with the actual sheet name
  
    // Get the data range
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
  
    // Iterate through rows starting from the 4th row (index 3)
    for (var i = 3; i < values.length; i++) {
      var row = values[i];
  
      // Extract relevant data
      var matricula = row[0];
      var aluno = row[1];
      var faltas = row[2];
      var p1 = row[3];
      var p2 = row[4];
      var p3 = row[5];
  
      // Calculate average
      var average = ((p1 + p2 + p3) / 10) / 3;
  
      // Check for Reproved by Nota
      if (average < 5) {
        sheet.getRange(i + 1, 7).setValue('Reprovado por Nota');
        sheet.getRange(i + 1, 8).setValue(0);
      } else if (average < 7) {
        // Check for Reproved by Falta
        var totalFaltas = 0.25 * 60; // 25% of total classes
        if (faltas > totalFaltas) {
          sheet.getRange(i + 1, 7).setValue('Reprovado por Falta');
          sheet.getRange(i + 1, 8).setValue(0);
        } else {
          // Check for Exame Final
          var naf = Math.ceil(10 - average);
          sheet.getRange(i + 1, 7).setValue('Exame Final');
          sheet.getRange(i + 1, 8).setValue(naf);
        }
      } else {
        // Aprovado
        var totalFaltas = 0.25 * 60; // 25% of total classes
        if (faltas > totalFaltas) {
          sheet.getRange(i + 1, 7).setValue('Reprovado por Falta');
          sheet.getRange(i + 1, 8).setValue(0);
        } else {
          sheet.getRange(i + 1, 7).setValue('Aprovado');
          sheet.getRange(i + 1, 8).setValue(0);
        }
      }
    }
  
    // Log completion outside of the loop
    Logger.log('Results calculated and updated successfully.');
  }
  