function ReplaceContractDetails() {
  var doctemplateid = "1ONMUpt9RZYXyYOrm-3wrz9LqGKsrH_5mrIYuiXxZ6I4";
  //var docfinalid = "1MbQG6c3TPc2czxRfegUbPvQXh-LTfWfBn4ruPxA-b_A";
  var wsid = "1uTSKh9ntrESqvoN8iQBPEns_9EAstCAc7BgVyLgopkw";
  //var wsurl = "docs.google.com/spreadsheets/d/1uTSKh9ntrESqvoN8iQBPEns_9EAstCAc7BgVyLgopkw/edit";
  // getting the id of the google sheet and google id template and final
  var rebateid = "1lqT1XopyUyVcQ3-3Y0g_jx0KZez792eqWx1-61X_1wE"
  var doctemplate = DocumentApp.openById(doctemplateid);
  //var docfinal = DocumentApp.openById(docfinalid);

  var ws = SpreadsheetApp.openById(wsid).getSheetByName("Contract Generation for Chinese Form")
 
  var rebatetemplate = DocumentApp.openById(rebateid)
  var data = ws.getRange(2,2,ws.getLastRow()-1,4).getValues();
  
  var datarange1 = ws.getRange(1, 12, 2, 2).getValues()
  var datarange2 = ws.getRange(1, 12, 3, 2).getValues()
  var datarange3 = ws.getRange(1, 12, 4, 2).getValues()
  var datarange4 = ws.getRange(1, 12, 5, 2).getValues()
  
  
  var templatepara = doctemplate.getBody().getParagraphs()
  var rebatepara = rebatetemplate.getBody().getParagraphs()
  //var templatetable = doctemplate.getBody().getTables()
  
  //var templatecopy = doctemplate.getBody().findElement(DocumentApp.ElementType.PARAGRAPH) //can research deeper into findelement
  //Logger.log(templatecopy)
  //var templatereplaced = templatecopy.replaceText("companyname", "LOL")
  //var newtemplatepara = templatereplaced.getParagraphs();
  //newtemplatepara.forEach(function(p){
    //docfinal.getBody().appendParagraph(p.copy())
  //});
  
  //Logger.log(doctemplate)
  //Logger.log(newtemplatepara)
  
  //docfinal.getBody().clear();
  

  
  //.appendTable(datarange1)
  var todaysdate = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd")
  
  for (var i = 0; i<ws.getLastRow()-1;i++){
    var statusplate = ws.getRange(i+2,1).getValues();
    var rebateplate = ws.getRange(i+2, 4).getValues();
    var nameplate = ws.getRange(i+2,2).getValues()
    
    if (statusplate == "" && rebateplate != "" && nameplate != ""){
      var finalcontractdoc = DocumentApp.create(data[i][0])
      templatepara.forEach(function(p){
        finalcontractdoc.getBody().appendParagraph(
        p
        .copy()
        .replaceText("companyname", data[i][0])
        .replaceText("address", data[i][1])
        .replaceText("date",todaysdate)
        )
      })
      
      
      var rebatecontractdoc = DocumentApp.create(data[i][0])
      var rebatebody = rebatecontractdoc.getBody()
      //var rebatetabletext = rebatebody.findText("<table>")
      //.getElement()
      //.getParent().asBody().appendTable(datarange1)
      
      var rangeBuilder = rebatecontractdoc.newRange()
      //rangeBuilder.addRange(datarange1)
      //var ind = rebatebody.getChildIndex(table)
      rebatepara.forEach(function(p){
        rebatecontractdoc.getBody().appendParagraph(
        p
        .copy()
        .replaceText("companyname", data[i][0])
        .replaceText("date",todaysdate)
        )
      })
      var rebatetabletext = rebatebody.findText("3、乙方将在季度结束后且收到甲方全部广告款项后").getElement().getParent()
      var tableDechildindex = rebatebody.getChildIndex(rebatetabletext)
      rebatebody.insertParagraph(tableDechildindex, "\n")
      if (ws.getRange(2, 11).getValues() == "y") {
        rebatebody.insertTable(tableDechildindex, datarange1)
      }
      if (ws.getRange(3, 11).getValues() == "y") {
        rebatebody.insertTable(tableDechildindex, datarange2)
      }
      if (ws.getRange(4, 11).getValues() == "y") {
        rebatebody.insertTable(tableDechildindex, datarange3)
      }
      if (ws.getRange(5, 11).getValues() == "y") {
        rebatebody.insertTable(tableDechildindex, datarange4)
      }

      var contracturl = finalcontractdoc.getUrl()
      var rebateurl = rebatecontractdoc.getUrl()
      ws.getRange(i + 2, 1).setValue("Contract Generated")
      ws.getRange(i+2, 6).setValue(contracturl)
      ws.getRange(i+2,7).setValue(rebateurl)
    }
      
    else if (statusplate == "" && nameplate !="") {
      var finalcontractdoc = DocumentApp.create(data[i][0])
      templatepara.forEach(function(p){
        finalcontractdoc.getBody().appendParagraph(
        p
        .copy()
        .replaceText("companyname",data[i][0])
        .replaceText("address",data[i][1])
        .replaceText("date", todaysdate)
        )
        })
      var contracturl = finalcontractdoc.getUrl()
      ws.getRange(i + 2, 1).setValue("Contract Generated")
      ws.getRange(i+2, 6).setValue(contracturl)
    };
      //templatetable.forEach(function(t){
        //docfinal.getBody().appendTable(t.copy()) took away this function becos i wanna show boss that this code works first. 31jul2019
      //});
    
      //i put this code at the bottom, so the table is at the bottom. How can I make it the way it is eh? can it be the append function
      //the video can arrange the list accordingly because the list is inside the paragraph
  }
}


   


  //data.forEach(function(r){
    //createMailMerge(r[0],r[1],templatepara,docfinal)
  //get template's paragraph. for each of the template's paragraph, we get the body of the docfinal and add the template'a para into it.
      
  
    
               
               
