      /**
       *Create a cool new menu
       **/
      function onOpen() {
        var ui = SpreadsheetApp.getUi();
        ui.
        createMenu('Custom Script')
          .addItem('Send Queue Notfication Email', 'queueNotifier')
          .addToUi();

        function queueNotifier() {
          // Send an email with the current X Assignment queue status.
        }
      }

      /**
       *Custom function that sends an email notification to a list of users
       *with the cases remaining in the X Assignment Queue by Due Date
       **/

      function queueNotifier() {
        var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Queue Notification');
        var data = sh.getSheetValues(1, 1, sh.getLastRow(), 2);
        var emailSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('emails');
        var emailData = emailSheet.getDataRange().getValues();
        var subject = 'X Assignment Queue Update'
        //var htmltable =[];

        var TABLEFORMAT = 'cellspacing="2" cellpadding="2" dir="ltr" border="1" style="width:100%;table-layout:fixed;font-size:10pt;font-family:arial,sans,sans-serif;border-collapse:collapse;border:1px solid #ccc;font-weight:normal;color:black;background-color:white;text-align:center;text-decoration:none;font-style:normal;'
        var htmltable = '<table ' + TABLEFORMAT + ' ">';

        for (row = 0; row < data.length; row++) {

          htmltable += '<tr>';

          for (col = 0; col < data[row].length; col++) {
            if (data[row][col] === "" || 0) {
              htmltable += '<td>' + 'None' + '</td>';
            } else
            if (row === 0) {
              htmltable += '<th>' + data[row][col] + '</th>';
            } else {
              htmltable += '<td>' + data[row][col] + '</td>';
            }
          }

          htmltable += '</tr>';
        }

        htmltable += '</table>';
        Logger.log(data);;
        for (var i in emailData) {
          var email = emailData[i].toString()
          MailApp.sendEmail(email, subject, '', {htmlBody: htmltable})
          Logger.log("Successful email sent to:", emailData[i])
        }
      }