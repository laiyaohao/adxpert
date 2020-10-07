//define global variable for identification of documents and cells
var docId = "";
var invoiceTemplate = DocumentApp.openById(docId);
// if change name, change the parenthesis inside of getSheetByName to the name of the sheet, adding single quotes around it.
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
var statusCells = sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues();
var invoiceNumberCells = sheet.getRange(2, 2, sheet.getLastRow()-1, 1).getValues();
//var dateCells = sheet.getRange(2, 2, sheet.getLastRow()-1, 1).getValues();
var clientNameCells = sheet.getRange(2,3,sheet.getLastRow()-1,1).getValues();
var clientAddressCells = sheet.getRange(2,4,sheet.getLastRow()-1,1).getValues();
var descriptionCells = sheet.getRange(2,5,sheet.getLastRow()-1,1).getValues();
var amountCells = sheet.getRange(2,6,sheet.getLastRow()-1,1).getValues();
var todaysDate = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd")


function generateInvoice() {
  var invoiceContent = invoiceTemplate.getBody().getTables();
  //var adxpertLogo = invoiceTemplate.getBody().getImages()[0].copy();
  var pictureIndex = (invoiceTemplate.getBody().getParagraphs().length) - 1;
  
  var filledRows = 0;
  
  for (var i = 0; i < sheet.getLastRow()-1; i++) {
    var statusCell = statusCells[i];
    var invoiceNumberCell = invoiceNumberCells[i];
    var clientNameCell = clientNameCells[i];
    var clientAddressCell = clientAddressCells[i];
    var descriptionCell = descriptionCells[i];
    var amountCell = amountCells[i]
    if (statusCell == '' && invoiceNumberCell != '') {
      var invoiceDoc = DocumentApp.create(invoiceNumberCells[i]);
      var adxpertLogo = invoiceTemplate.getBody().getParagraphs()[pictureIndex].copy()
      invoiceContent.forEach(function(p){
        invoiceDoc.getBody().appendTable(
        p
        .copy()
        .replaceText("invoice_number", invoiceNumberCell)
        .replaceText("date", todaysDate)
        .replaceText("client_name",clientNameCell)
        .replaceText("client_address",clientAddressCell)
        .replaceText("description",descriptionCell)
        .replaceText("amount1",amountCell)
        .replaceText("amount2",amountCell)
        )
        
      })
      invoiceDoc.getBody().appendParagraph(adxpertLogo); //sometimes can work, sometimes cannot work for images
      sheet.getRange(i+2,1).setValue("Invoice Generated on " + Utilities.formatDate(new Date(), "GMT+8","ddMMyy"))
    }
  }
}
function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('Generate Invoice').addItem('Generate Invoice', 'generateInvoice').addToUi()
}
