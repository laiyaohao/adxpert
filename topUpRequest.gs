var ss = SpreadsheetApp.getActiveSpreadsheet()
ss.setActiveSheet(ss.getSheetByName('Top Up Function'))
var sheet = SpreadsheetApp.getActiveSheet()
var statuscells = sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues()
var resellercells = sheet.getRange(2, 2, sheet.getLastRow()-1, 1).getValues()
var idcells = sheet.getRange(2, 3, sheet.getLastRow()-1, 1).getValues()
var moneycells = sheet.getRange(2,4,sheet.getLastRow()-1,1).getValues()
var decisioncells = sheet.getRange(2,5,sheet.getLastRow(),1).getValues()
var data = sheet.getRange(2,3,sheet.getLastRow()-1,3).getValues()
var header = sheet.getRange(1, 3, 1, 3).getValues()[0]
//var emailsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet3')
var emailaddress1 = sheet.getRange(1,7).getValues()
var emailaddress2 = sheet.getRange(1,8).getValues()
var cheetahemailaddresses = emailaddress1 + ',' + emailaddress2
var emailaddress3 = sheet.getRange(2,7).getValues()
var emailaddress4 = sheet.getRange(2,8).getValues()
var emailaddress5 = sheet.getRange(2,9).getValues()
var pandaemailaddresses = emailaddress3+','+emailaddress4 + ',' + emailaddress5

function firstTopUpRequest() {
  var countofdetail = 0;
  var i = 0;
  var u = 0; //number of rows that is done. did this so that i can count which row to start the extraction of data from
  var cheetahmessage = []
  var pandamessage = []
  // want to check which row the details are on. cos above rows may be done le, so need to check which rows to start from
  /*for (; i<sheet.getLastRow()-1; i++) {
    if (statuscells[i] !=''){
    u++
    }
    }*/
    // u is the row number.
  for (; i<sheet.getLastRow()-1; i++) {
  
  // if reseller is cheetah:
  
    if (statuscells[i] == '' && idcells[i] != '' && moneycells[i] != '' && decisioncells[i] != '' && resellercells[i] == 'cheetah') {
    cheetahmessage.push(data[i])
    }
    if (statuscells[i] == '' && idcells[i] != '' && moneycells[i] != '' && decisioncells[i] != '' && resellercells[i] == 'panda') {
    pandamessage.push(data[i])
      //countofdetail ++ //check if there is a need to send excel or i can just send an email with the request inside
      
      }
      }
   //f (cheetahmessage != undefined) {
   //if there are 3 or more rows, we attach excel sheet in the email.
    if (cheetahmessage.length >= 3) {
     var newsheet = SpreadsheetApp.create('充值模板_' + Utilities.formatDate(new Date(), "GMT+8", "ddMMyy")).getActiveSheet()
      newsheet.appendRow(header)
      for (; u < cheetahmessage.length; u++) {
      newsheet.appendRow(cheetahmessage[u])
      }
      var newexcel = newsheet.getParent()
      var url = "https://docs.google.com/feeds/download/spreadsheets/Export?key=" + newexcel.getId() + "&exportFormat=xlsx";
      var params = {
        method      : "get",
        headers     : {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
        muteHttpExceptions: true
      };
       var blob1 = UrlFetchApp.fetch(url,params).getBlob()
       newsheet.setName('充值模板_' + Utilities.formatDate(new Date(), "GMT+8", "ddMMyy"))
       blob1.setName(newsheet.getName() + ".xlsx");
       var emailsub = 'adxpert-Facebook-账户管理-' + Utilities.formatDate(new Date(), "GMT+8","ddMMyy")
       MailApp.sendEmail(cheetahemailaddresses, emailsub, "燕星，麻烦转款/充值/扣减", {attachments: blob1})
      for (var q = 0; q < sheet.getLastRow()-1; q++) {
      if (statuscells[q] == '' && idcells[q] != '' && moneycells[q] != '' && decisioncells[q] != '' && resellercells[q] == 'cheetah') {
      sheet.getRange(q+2, 1).setValue('top up sent to cheetah on ' + Utilities.formatDate(new Date(), "GMT+8","ddMMyy"))
      }
      }
    }
    
    //if it is less than 3 rows, ie 2 rows or less, we send email with text
    if (cheetahmessage.length < 3 && cheetahmessage.length > 0) {
      var emailsub = 'adxpert-Facebook-账户管理-' + Utilities.formatDate(new Date(), "GMT+8","ddMMyy")
      var decisionarray = []
      var lesserthan2 = []
      for (; u < cheetahmessage.length; u++) {
        decisionarray.push(cheetahmessage[u][2])
      /*if (decisioncells[u] == '2') {
      var decision = '扣减'
      cheetahmessage.push('麻烦为' + acctnumber + '扣减' + moneycells[u])
      }*/
      
      }
      for (; countofdetail < cheetahmessage.length; countofdetail ++) {
      var acctnummber = cheetahmessage[countofdetail][0]
      var amt2transfer = cheetahmessage[countofdetail][1]
      if (decisionarray[countofdetail] == '1') {
      lesserthan2.push('麻烦为' + cheetahmessage[countofdetail][0] + '充值' + cheetahmessage[countofdetail][1])
      
      }
      if (decisionarray[countofdetail] == '2') {
      lesserthan2.push('麻烦为' + cheetahmessage[countofdetail][0] + '扣减' + cheetahmessage[countofdetail][1])
      
      }
      }
      lesserthan2.join()
      var emailsub = 'adxpert-Facebook-账户管理-' + Utilities.formatDate(new Date(), "GMT+8","ddMMyy")
      MailApp.sendEmail(cheetahemailaddresses, emailsub, lesserthan2)
      for (var q = 0; q < sheet.getLastRow()-1; q++) {
      if (statuscells[q] == '' && idcells[q] != '' && moneycells[q] != '' && decisioncells[q] != '' && resellercells[q] == 'cheetah') {
      sheet.getRange(q+2, 1).setValue('top up sent to cheetah on ' + Utilities.formatDate(new Date(), "GMT+8","ddMMyy"))
      }
      }
    }
    
    
    //panda doesnt need any excel sheet, so it is okay to just send email
   if (pandamessage.length > 0) {
    var emailsub = 'adxpert-Facebook-账户管理' + Utilities.formatDate(new Date(), "GMT+8","ddMMyy")
      var decisionarray = []
      var lesserthan2 = []
      for (; u < pandamessage.length; u++) {
        decisionarray.push(pandamessage[u][2])
      /*if (decisioncells[u] == '2') {
      var decision = '扣减'
      cheetahmessage.push('麻烦为' + acctnumber + '扣减' + moneycells[u])
      }*/
      
      }
      for (; countofdetail < pandamessage.length; countofdetail ++) {
      if (decisionarray[countofdetail] == '1') {
      lesserthan2.push('麻烦为' + pandamessage[countofdetail][0] + '充值' + pandamessage[countofdetail][1])
      
      }
      if (decisionarray[countofdetail] == '2') {
      lesserthan2.push('麻烦为' + pandamessage[countofdetail][0] + '扣减' + pandamessage[countofdetail][1])
      
      }
      }
      lesserthan2.join()
      var emailsub = 'adxpert-Pandamobo-Facebook-账户管理-' + Utilities.formatDate(new Date(), "GMT+8","ddMMyy")
      MailApp.sendEmail(pandaemailaddresses, emailsub, lesserthan2)
      for (var q = 0; q < sheet.getLastRow()-1; q++) {
      if (statuscells[q] == '' && idcells[q] != '' && moneycells[q] != '' && decisioncells[q] != '' && resellercells[q] == 'panda') {
      sheet.getRange(q+2, 1).setValue('top up sent to panda on ' + Utilities.formatDate(new Date(), "GMT+8","ddMMyy"))
      }
      }
   }
   
   
   }
   
function subsequentTopUp() {
  var countofdetail = 0;
  var i = 0;
  var u = 0; //number of rows that is done. did this so that i can count which row to start the extraction of data from
  var cheetahmessage = []
  var pandamessage = []
  // want to check which row the details are on. cos above rows may be done le, so need to check which rows to start from
  /*for (; i<sheet.getLastRow()-1; i++) {
    if (statuscells[i] !=''){
    u++
    }
    }*/
    // u is the row number.
  for (; i<sheet.getLastRow()-1; i++) {
  
  // if reseller is cheetah:
  
    if (statuscells[i] == '' && idcells[i] != '' && moneycells[i] != '' && decisioncells[i] != '' && resellercells[i] == 'cheetah') {
    cheetahmessage.push(data[i])
    }
    if (statuscells[i] == '' && idcells[i] != '' && moneycells[i] != '' && decisioncells[i] != '' && resellercells[i] == 'panda') {
    pandamessage.push(data[i])
      //countofdetail ++ //check if there is a need to send excel or i can just send an email with the request inside
      
      }
      }
   //f (cheetahmessage != undefined) {
   //if there are 3 or more rows, we attach excel sheet in the email.
    if (cheetahmessage.length >= 3) {
     var newsheet = SpreadsheetApp.create('充值模板_' + Utilities.formatDate(new Date(), "GMT+8", "ddMMyy")).getActiveSheet()
      newsheet.appendRow(header)
      for (; u < cheetahmessage.length; u++) {
      newsheet.appendRow(cheetahmessage[u])
      }
      var newexcel = newsheet.getParent()
      var url = "https://docs.google.com/feeds/download/spreadsheets/Export?key=" + newexcel.getId() + "&exportFormat=xlsx";
      var params = {
        method      : "get",
        headers     : {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
        muteHttpExceptions: true
      };
       var blob1 = UrlFetchApp.fetch(url,params).getBlob()
       newsheet.setName('充值模板_' + Utilities.formatDate(new Date(), "GMT+8", "ddMMyy"))
       blob1.setName(newsheet.getName() + ".xlsx");
       //var emailsub = 'adxpert-Facebook-账户管理-' + Utilities.formatDate(new Date(), "GMT+8","ddMMyy")
       //MailApp.sendEmail(cheetahemailaddresses, emailsub, "燕星，麻烦转款/充值/扣减", {attachments: blob1})
       
       var gmailquary = GmailApp.search('账户管理 ' + Utilities.formatDate(new Date(), "GMT+8", "ddMMyy") + ' haoyanxing')
       gmailquary[0].replyAll("燕星，麻烦转款/充值/扣减", {attachments: blob1})
       
       //filling the cells in excel sheet to show the top up is done
      for (var q = 0; q < sheet.getLastRow()-1; q++) {
      if (statuscells[q] == '' && idcells[q] != '' && moneycells[q] != '' && decisioncells[q] != '' && resellercells[q] == 'cheetah') {
      sheet.getRange(q+2, 1).setValue('top up sent to cheetah on ' + Utilities.formatDate(new Date(), "GMT+8","ddMMyy"))
      }
      }
    }
    
    //if it is less than 3 rows, ie 2 rows or less, we send email with text
    if (cheetahmessage.length < 3 && cheetahmessage.length > 0) {
      //var emailsub = 'adxpert-Facebook-账户管理-' + Utilities.formatDate(new Date(), "GMT+8","ddMMyy")
      var decisionarray = []
      var lesserthan2 = []
      for (; u < cheetahmessage.length; u++) {
        decisionarray.push(cheetahmessage[u][2])
      /*if (decisioncells[u] == '2') {
      var decision = '扣减'
      cheetahmessage.push('麻烦为' + acctnumber + '扣减' + moneycells[u])
      }*/
      
      }
      for (; countofdetail < cheetahmessage.length; countofdetail ++) {
      var acctnummber = cheetahmessage[countofdetail][0]
      var amt2transfer = cheetahmessage[countofdetail][1]
      if (decisionarray[countofdetail] == '1') {
      lesserthan2.push('麻烦为' + cheetahmessage[countofdetail][0] + '充值' + cheetahmessage[countofdetail][1])
      
      }
      if (decisionarray[countofdetail] == '2') {
      lesserthan2.push('麻烦为' + cheetahmessage[countofdetail][0] + '扣减' + cheetahmessage[countofdetail][1])
      
      }
      }
      lesserthan2.join()
      //var emailsub = 'adxpert-Facebook-账户管理-' + Utilities.formatDate(new Date(), "GMT+8","ddMMyy")
      //MailApp.sendEmail(cheetahemailaddresses, emailsub, lesserthan2)
      
      var gmailquary = GmailApp.search('账户管理 ' + Utilities.formatDate(new Date(), "GMT+8", "ddMMyy") + ' haoyanxing')
       gmailquary[0]
       .replyAll(lesserthan2)
      
      //
      for (var q = 0; q < sheet.getLastRow()-1; q++) {
      if (statuscells[q] == '' && idcells[q] != '' && moneycells[q] != '' && decisioncells[q] != '' && resellercells[q] == 'cheetah') {
      sheet.getRange(q+2, 1).setValue('top up sent to cheetah on ' + Utilities.formatDate(new Date(), "GMT+8","ddMMyy"))
      }
      }
    }
    
    //panda doesnt need any excel sheet, so it is okay to just send email
   if (pandamessage.length > 0) {
    //var emailsub = 'adxpert-Pandamobo-Facebook-账户管理' + Utilities.formatDate(new Date(), "GMT+8","ddMMyy")
      var decisionarray = []
      var lesserthan2 = []
      for (; u < pandamessage.length; u++) {
        decisionarray.push(pandamessage[u][2])
      /*if (decisioncells[u] == '2') {
      var decision = '扣减'
      cheetahmessage.push('麻烦为' + acctnumber + '扣减' + moneycells[u])
      }*/
      
      }
      for (; countofdetail < pandamessage.length; countofdetail ++) {
      if (decisionarray[countofdetail] == '1') {
      lesserthan2.push('麻烦为' + pandamessage[countofdetail][0] + '充值' + pandamessage[countofdetail][1])
      
      }
      if (decisionarray[countofdetail] == '2') {
      lesserthan2.push('麻烦为' + pandamessage[countofdetail][0] + '扣减' + pandamessage[countofdetail][1])
      
      }
      }
      lesserthan2.join()
      var emailsub = 'adxpert-Pandamobo-Facebook-账户管理-' + Utilities.formatDate(new Date(), "GMT+8","ddMMyy")
      MailApp.sendEmail(pandaemailaddresses, emailsub, lesserthan2)
      
      //filling the cells in excel sheet to show the top up is done
      for (var q = 0; q < sheet.getLastRow()-1; q++) {
      if (statuscells[q] == '' && idcells[q] != '' && moneycells[q] != '' && decisioncells[q] != '' && resellercells[q] == 'panda') {
      sheet.getRange(q+2, 1).setValue('top up sent to panda on ' + Utilities.formatDate(new Date(), "GMT+8","ddMMyy"))
      }
      }
   }
}
   
function sendTopUpRequest() {
//check if there is an top up request today. If have reply to the email, instead of creating a new email thread. if dont have, create a new one
  var gmailquary = GmailApp.search('账户管理 ' + Utilities.formatDate(new Date(), "GMT+8", "ddMMyy"))
  
  //.getMessages()[1].getBody() //[0] is counting from the top
  var gmailquarylength = gmailquary.length
  /*Utilities.formatDate(new Date(), "GMT+8", "ddMMyy")*/
  if (gmailquary.length == 0) {
    firstTopUpRequest()
    }
    else {
      subsequentTopUp()
      }
}
