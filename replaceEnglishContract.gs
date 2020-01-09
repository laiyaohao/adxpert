function ReplaceContractDetails() {
  var doctemplateid = "1THon4uQ8vKBWSZzn5F1ciXE2vLsQFKFGUFRxaQ-ZfK8";
  //var docfinalid = "1MbQG6c3TPc2czxRfegUbPvQXh-LTfWfBn4ruPxA-b_A";
  var wsid = "1uTSKh9ntrESqvoN8iQBPEns_9EAstCAc7BgVyLgopkw";
  //var wsurl = "docs.google.com/spreadsheets/d/1uTSKh9ntrESqvoN8iQBPEns_9EAstCAc7BgVyLgopkw/edit";
  // getting the id of the google sheet and google id template and final

  var doctemplate = DocumentApp.openById(doctemplateid);
  //var docfinal = DocumentApp.openById(docfinalid);

  var ws = SpreadsheetApp.openById(wsid).getSheetByName("Contract Generation for English Form")
  var templatepara = doctemplate.getBody().getParagraphs()
  
  var data = ws.getRange(2, 2, ws.getLastRow()-1, 3).getValues()
  
  var datarange1 = ws.getRange(1, 7, 2, 2).getValues()
  var datarange2 = ws.getRange(1, 7, 3, 2).getValues()
  var datarange3 = ws.getRange(1, 7, 4, 2).getValues()
  var datarange4 = ws.getRange(1, 7, 5, 2).getValues()
  var datarange5 = ws.getRange(1, 7, 6, 2).getValues()
  
  var templatepara = doctemplate.getBody().getParagraphs()
  
  var todaysdate = Utilities.formatDate(new Date(), "GMT+8", "MMMMM dd, yyyy")
  
  for (var i = 0; i<ws.getLastRow()-1;i++){
    var statusplate = ws.getRange(i+2,1).getValues();
    var rebateplate = ws.getRange(i+2, 4).getValues();
    var nameplate = ws.getRange(i+2,2).getValues()
    
    if (statusplate == "" && nameplate != "") {
      var finalcontractdoc = DocumentApp.create(data[i][0])
      if (rebateplate == "y") {
        var finalcontractbody = finalcontractdoc.getBody()
        templatepara.forEach(function(p){
          finalcontractdoc.getBody().appendParagraph(
          p
          .copy()
          .replaceText("companyname", data[i][0])
          .replaceText("contactperson", data[i][1])
          .replaceText("<<date>>", todaysdate)
          )})
        var finalcontractimportanttext = finalcontractbody.findText("2.	Length of Contract").getElement().getParent()
        var tablechildindex = finalcontractbody.getChildIndex(finalcontractimportanttext)
        finalcontractbody.insertParagraph(tablechildindex, "\n")
        if (ws.getRange(2, 6).getValues() == "y") {
          finalcontractbody.insertTable(tablechildindex, datarange1)
        }
        if (ws.getRange(3, 6).getValues() == "y") {
          finalcontractbody.insertTable(tablechildindex, datarange2)
        }
        if (ws.getRange(4, 6).getValues() == "y") {
          finalcontractbody.insertTable(tablechildindex, datarange3)
        }
        if (ws.getRange(5, 6).getValues() == "y") {
          finalcontractbody.insertTable(tablechildindex, datarange4)
        }
        if (ws.getRange(6, 6).getValues() == "y") {
          finalcontractbody.insertTable(tablechildindex, datarange5)
        }
        finalcontractbody.insertParagraph(tablechildindex, "1.4      Agency will provide client 0 – 3% rebate based on cumulative ad spend per quarter based on the following tiering" + "\n")
        //finalcontractbody.insertParagraph(finalcontractbody.getChildIndex(finalcontractbody.findText("Cumulative Ad spent per quarter").getElement().getParent()), "\n")
        var contracturl = finalcontractdoc.getUrl()
        ws.getRange(i + 2, 1).setValue("Contract Generated")
        ws.getRange(i+2, 5).setValue(contracturl)
        }
      else if (rebateplate == "") {
        var finalcontractbody = finalcontractdoc.getBody()
        templatepara.forEach(function(p){
          finalcontractdoc.getBody().appendParagraph(
          p
          .copy()
          .replaceText("companyname", data[i][0])
          .replaceText("contactperson", data[i][1])
          .replaceText("<<date>>", todaysdate)
          )})
        var contracturl = finalcontractdoc.getUrl()
        ws.getRange(i+2,1).setValue("Contract Generated")
        ws.getRange(i+2,5).setValue(contracturl)
      }
     }
     }
     }
