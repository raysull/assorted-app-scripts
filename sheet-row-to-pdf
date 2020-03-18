/* heavily inspired by andrewroberts' script

PDF Creator - Email all responses
=================================

When you click "Create PDF > Create a PDF for each row" this script 
constructs a PDF for each row in the attached GSheet. The value in the 
"File Name" column is used to name the file and - if there is a 
value - it is emailed to the recipient in the "Email" column.

Demo sheet with script attached: https://goo.gl/sf02mK

*/

// Config
// ------


// 1. Create a GDoc template and put the ID here

var TEMPLATE_ID = '---- UPDATE ME -----'

// 2. You can specify a name for the new PDF file here, or leave empty to use the 
// name of the template or specify the file name in the sheet

var PDF_FILE_NAME = ''

// 3. If an email address is specified you can email the PDF

var EMAIL_SUBJECT = 'The email subject ---- UPDATE ME -----'
var EMAIL_BODY = 'The email body ------ UPDATE ME ---------'

// 4. If a folder ID is specified here this is where the PDFs will be located

var RESULTS_FOLDER_ID = ''

// Constants
// ---------

// You can pull out specific columns values 
var FILE_NAME_COLUMN_NAME = 'Document Name'
var EMAIL_COLUMN_NAME = 'Email'

// The format used for any dates 
var DATE_FORMAT = 'yyyy/MM/dd';

/**
 * Eventhandler for spreadsheet opening - add a menu.
 */

function onOpen() {

  SpreadsheetApp
    .getUi()
    .createMenu('Create PDFs')
    .addItem('Create a PDF for each row', 'createPdfs')
    .addToUi()

} // onOpen()

/**  
 * Take the fields from each row in the active sheet
 * and, using a Google Doc template, create a PDF doc with these
 * fields replacing the keys in the template. The keys are identified
 * by having a % either side, e.g. %Name%.
 */

function createPdfs() {

  var ui = SpreadsheetApp.getUi()

  if (TEMPLATE_ID === '') {    
    ui.alert('TEMPLATE_ID needs to be defined in code.gs')
    return
  }

  // Set up the docs and the spreadsheet access

  var templateFile = DriveApp.getFileById(TEMPLATE_ID)
  var activeSheet = SpreadsheetApp.getActiveSheet()
  var allRows = activeSheet.getDataRange().getValues()
  var headerRow = allRows.shift()

  // Create a PDF for each row

  allRows.forEach(function(row) {
  
    createPdf(templateFile, headerRow, row)
    
    // Private Function
    // ----------------
  
    /**
     * Create a PDF
     *
     * @param {File} templateFile
     * @param {Array} headerRow
     * @param {Array} activeRow
     */
  
    function createPdf(templateFile, headerRow, activeRow) {
      
      var headerValue
      var activeCell
      var ID = null
      var recipient = null
      var copyFile
      var numberOfColumns = headerRow.length
      var copyFile = templateFile.makeCopy()      
      var copyId = copyFile.getId()
      var copyDoc = DocumentApp.openById(copyId)
      var copyBody = copyDoc.getActiveSection()
           
      // Replace the keys with the spreadsheet values and look for a couple
      // of specific values
     
      for (var columnIndex = 0; columnIndex < numberOfColumns; columnIndex++) {
        
        headerValue = headerRow[columnIndex]
        activeCell = activeRow[columnIndex]
        activeCell = formatCell(activeCell);
                
        copyBody.replaceText('%' + headerValue + '%', activeCell)
        
        if (headerValue === FILE_NAME_COLUMN_NAME) {
        
          ID = activeCell
          
        } else if (headerValue === EMAIL_COLUMN_NAME) {
        
          recipient = activeCell
        }
      }
      
      // Create the PDF file
        
      copyDoc.saveAndClose()
      var newFile = DriveApp.createFile(copyFile.getAs('application/pdf'))  
      copyFile.setTrashed(true)
    
      // Rename the new PDF file
    
      if (PDF_FILE_NAME !== '') {
      
        newFile.setName(PDF_FILE_NAME)
        
      } else if (ID !== null){
    
        newFile.setName(ID)
      }
      
      // Put the new PDF file into the results folder
      
      if (RESULTS_FOLDER_ID !== '') {
      
        DriveApp.getFolderById(RESULTS_FOLDER_ID).addFile(newFile)
        DriveApp.removeFile(newFile)
      }

      // Email the new PDF

      if (recipient !== null) {
      
        MailApp.sendEmail(
          recipient, 
          EMAIL_SUBJECT, 
          EMAIL_BODY,
          {attachments: [newFile]})
      }
    
    } // createPdfs.createPdf()

  })

  ui.alert('New PDF files created')

  return
  
  // Private Functions
  // -----------------
  
  /**
  * Format the cell's value
  *
  * @param {Object} value
  *
  * @return {Object} value
  */
  
  function formatCell(value) {
    
    var newValue = value;
    
    if (newValue instanceof Date) {
      
      newValue = Utilities.formatDate(
        value, 
        Session.getScriptTimeZone(), 
        DATE_FORMAT);
        
    } else if (typeof value === 'number') {
    
      newValue = Math.round(value * 100) / 100
    }
    
    return newValue;
        
  } // createPdf.formatCell()
  
} // createPdfs()
