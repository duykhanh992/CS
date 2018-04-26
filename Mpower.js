function doPost(e) {

    try {
        Logger.log(e); // the Google Script version of console.log see: Class Logger
        record_data(e); //to enter data to excel sheet
        //SendEmailAfterRecord(e); //
        return HtmlService.createHtmlOutput('<p><strong>OPT Application submitted successfully!</strong></p><p><strong>The request will take up to 1 business day to process. We will contact you <b>only</b> in case of any issues or concerns.&nbsp;</strong></p><p><strong style="color: rgb(255, 0, 0);">Note: Email/Call</strong><span style="color: rgb(255, 0, 0);">&nbsp;</span><strong style="color: rgb(255, 0, 0);">ISS</strong><span style="color: rgb(255, 0, 0);">&nbsp;</span><strong style="color: rgb(255, 0, 0);">only if you don&#39;t receive a response after 1 business day from the day of your application submission. Any violation in this may delay the processing further.</strong>NOTE: Only in case you do not get an email with your quiz results within one business day please send an email to tgyan@okstate.edu (Please do this <strong>only</strong> if you do not get your quiz results)</p><p>&nbsp;</p>');
        // return e;
    } catch (error) { // if error return this
        Logger.log(error);
        return ContentService
            .createTextOutput(JSON.stringify({
                "result": "error",
                "error": e
            }))
            .setMimeType(ContentService.MimeType.JSON);
    }
}

function record_data(e) {
    //Code block - to insert data from html page to google spreadsheet
    Logger.log(JSON.stringify(e)); // log the POST data in case we need to debug it
    try {

        var doc = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = doc.getSheetByName('Sheet1'); // select the responses sheet
        var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        var nextRow = sheet.getLastRow() + 1; // get next row
        var row = [new Date()]; // first element in the row should always be a timestamp

        // loop through the header columns
        for (var i = 1; i < headers.length; i++) { // start at 1 to avoid Timestamp column
            if (headers[i].length > 0) {
                row.push(e.parameter[headers[i]]); // add data to row
            }
        }
        //end of for loop for saving data
          var dateOfSubmission = String(row[0]);
            var lastName = String(row[1]);
            var firstName = String(row[2]);
            var middleName = String(row[3]);
            var birthDate = String(row[4]);
            var employment = String(row[5]);
            var income = String(row[6]);
            var email = String(row[7]);
            var phone = String(row[8]);
            var address1 = String(row[9]);
            var address2  = String(row[10]);
            var city = String(row[11]);
            var state = String(row[12]);
            var zipcode = String(row[13]);
            var country = String(row[14]);

            
            var excel = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1kRWJWuGbG_ruZlyjBlKTKemNYko9IuaRmM7rOk-xtfM/edit#gid=0");
            // var results = excel.getActiveSpreadSheet();
            var resultSheet = excel.getSheetByName('Sheet1');
            var startRow = 2; // First row of data to process
            var numRows = 10000; // Number of rows to process
            // Fetch the range of cells A2:D5
            var dataRange = resultSheet.getRange(startRow, 1, numRows, 1000)

            // Fetch values for each row in the Range.
            var data = dataRange.getValues();
    }
      catch (error) {
        Logger.log(error);
        return ContentService
            .createTextOutput(JSON.stringify({
                "result": "error",
                "error": e
            }))
            .setMimeType(ContentService.MimeType.JSON);


    } 
    finally {

        return;
    }
}
