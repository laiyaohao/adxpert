var ss = SpreadsheetApp.getActiveSpreadsheet()
ss.setActiveSheet(ss.getSheetByName('Top Up Function'))
var sheet = SpreadsheetApp.getActiveSheet()

// the first line basically means getting the get the active spreadsheet, which is the current spreadsheet
// as fpr the second line, i set the top up function as the active sheet
// and the third line means getting the active sheet, which is the top up function sheet that i have just set.

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
var emailaddress3 = sheet.getRange(1,9).getValues()
var cheetahemailaddresses = emailaddress1 + ',' + emailaddress2 + ',' + emailaddress3
var emailaddress4 = sheet.getRange(2,7).getValues()
var emailaddress5 = sheet.getRange(2,8).getValues()
var emailaddress6 = sheet.getRange(2,9).getValues()
var pandaemailaddresses = emailaddress3+','+emailaddress4 + ',' + emailaddress5 + ',' + emailaddress6
var yinolinkemailaddresses = sheet.getRange(3,7).getValues()

//the above lines of code basically means getting the values from the cells i want and storing it into variables. for the api for all these, refer to:
//https://developers.google.com/apps-script/reference/spreadsheet/sheet.html

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
    var yinolinkmessage = []
    
    //creating 3 variables for the 3 resellers and equating them to empty arrays
    //the idea is to fill up the arrays with the code below and send them to respective resellers
  for (; i<sheet.getLastRow()-1; i++) {
  
  // if reseller is cheetah:
  
    if (statuscells[i] == '' && idcells[i] != '' && moneycells[i] != '' && decisioncells[i] != '' && resellercells[i] == 'cheetah') {
    cheetahmessage.push(data[i]) //push means getting the data into the array that i have created.
    }
    
    //if reseller is panda:
    if (statuscells[i] == '' && idcells[i] != '' && moneycells[i] != '' && decisioncells[i] != '' && resellercells[i] == 'panda') {
    pandamessage.push(data[i])
      
      
      }
      //if reseller is yinolink:
      if (statuscells[i] == '' && idcells[i] != '' && moneycells[i] != '' && decisioncells[i] != '' && resellercells[i] == 'yinolink') {
      yinolinkmessage.push(data[i])
      }
      }
   
   //if there are 3 or more rows, we attach excel sheet in the email.
    if (cheetahmessage.length >= 3) {
     
     setTimeout(cheetahFirstTopUpUsingExcel(), 1000)
     //setTimeout is a default function that can be used, so that i can delay the sending of the email by 1000 milliseconds.
     //reason why i want to do that is explained in the notion documentation.
    }
    
    //if it is less than 3 rows, ie 2 rows or less, we send email with text
    if (cheetahmessage.length < 3 && cheetahmessage.length > 0) {
      
      setTimeout(cheetahFirstTopUpEmail(), 1000)
    }
    
    
    //panda doesnt need any excel sheet, so it is okay to just send email
   if (pandamessage.length > 0) {
    setTimeout(pandaTopUpRequest(), 1000)
   }
   //same as panda
   if (yinolinkmessage.length > 0) {
    setTimeout(yinolinkTopUpRequest(), 1000)
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
    var yinolinkmessage = []
  for (; i<sheet.getLastRow()-1; i++) {
  
  // if reseller is cheetah:
  
    if (statuscells[i] == '' && idcells[i] != '' && moneycells[i] != '' && decisioncells[i] != '' && resellercells[i] == 'cheetah') {
    cheetahmessage.push(data[i])
    }
    if (statuscells[i] == '' && idcells[i] != '' && moneycells[i] != '' && decisioncells[i] != '' && resellercells[i] == 'panda') {
    pandamessage.push(data[i])
      //countofdetail ++ //check if there is a need to send excel or i can just send an email with the request inside
      
      }
      if (statuscells[i] == '' && idcells[i] != '' && moneycells[i] != '' && decisioncells[i] != '' && resellercells[i] == 'yinolink') {
      yinolinkmessage.push(data[i])
      }
      }
   
   //if there are 3 or more rows, we attach excel sheet in the email.
    if (cheetahmessage.length >= 3) {
     setTimeout(cheetahReplyExcel(), 1000)
    }
    
    //if it is less than 3 rows, ie 2 rows or less, we send email with text
    if (cheetahmessage.length < 3 && cheetahmessage.length > 0) {
      setTimeout(cheetahReplyEmail(), 1000)
    }
    
    //panda doesnt need any excel sheet, so it is okay to just send email
   if (pandamessage.length > 0) {
    setTimeout(pandaTopUpRequest(), 1000)
   }
   //same as panda
   if (yinolinkmessage.length > 0) {
    setTimeout(yinolinkTopUpRequest(), 1000)
   }
}
   
function sendTopUpRequest() {
//check if there is an top up request today. If have reply to the email, instead of creating a new email thread. if dont have, create a new one
  var gmailquary1 = GmailApp.search('账户管理 ' + Utilities.formatDate(new Date(), "GMT+8", "ddMMyy") + '郝燕星' + '猎豹移动')
  
  
  if (gmailquary1.length == 0) {
    firstTopUpRequest()
    }
    else {
      subsequentTopUp()
      }
}

function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('Top up').addItem('Send Top Up Request', 'sendTopUpRequest').addToUi()
}

function pandaTopUpRequest() {
  var decisionarray = [] //creating an empty array for the decision of top up or withdraw
      var pandarealmessage = [] //creating an array for the message to be sent to panda
      for (; u < pandamessage.length; u++) {
        decisionarray.push(pandamessage[u][2])
      } // append the decision number which are full of '1's and '2's
      for (; countofdetail < pandamessage.length; countofdetail ++) { //for each top up / withdrawal for panda
      if (decisionarray[countofdetail] == '1') {
      pandarealmessage.push('麻烦为' + pandamessage[countofdetail][0] + '充值' + pandamessage[countofdetail][1])
      // if its '1', append the message such that its 充值
      }
      if (decisionarray[countofdetail] == '2') {
      pandarealmessage.push('麻烦为' + pandamessage[countofdetail][0] + '扣减' + pandamessage[countofdetail][1])
      //if its '2', append the message such that its withdrawal
      }
      }
      pandarealmessage.join() // join all the message together
      var emailsub = 'adxpert-Pandamobo-Facebook-账户管理-' + Utilities.formatDate(new Date(), "GMT+8","ddMMyy")
      MailApp.sendEmail(pandaemailaddresses, emailsub, pandarealmessage)
      
      //filling the cells in excel sheet to show the top up is done
      for (var q = 0; q < sheet.getLastRow()-1; q++) {
      if (statuscells[q] == '' && idcells[q] != '' && moneycells[q] != '' && decisioncells[q] != '' && resellercells[q] == 'panda') {
      sheet.getRange(q+2, 1).setValue('top up sent to panda on ' + Utilities.formatDate(new Date(), "GMT+8","ddMMyy"))
      }
      }
}

function yinolinkTopUpRequest() { //all of them same as panda

      var decisionarray = []
      var yinolinkrealmessage = []
      for (; u < yinolinkmessage.length; u++) {
        decisionarray.push(yinolinkmessage[u][2])
      
      
      }
      for (; countofdetail < yinolinkmessage.length; countofdetail ++) {
      if (decisionarray[countofdetail] == '1') {
      yinolinkrealmessage.push('麻烦为' + yinolinkmessage[countofdetail][0] + '充值' + yinolinkmessage[countofdetail][1])
      
      }
      if (decisionarray[countofdetail] == '2') {
      yinolinkrealmessage.push('麻烦为' + yinolinkmessage[countofdetail][0] + '扣减' + yinolinkmessage[countofdetail][1])
      
      }
      }
      yinolinkrealmessage.join()
      var emailsub = 'adxpert-Yinolink-账户管理-' + Utilities.formatDate(new Date(), "GMT+8","ddMMyy")
      MailApp.sendEmail(yinolinkemailaddresses, emailsub, yinolinkrealmessage)
      for (var q = 0; q < sheet.getLastRow()-1; q++) {
      if (statuscells[q] == '' && idcells[q] != '' && moneycells[q] != '' && decisioncells[q] != '' && resellercells[q] == 'yinolink') {
      sheet.getRange(q+2, 1).setValue('top up sent to yinolink on ' + Utilities.formatDate(new Date(), "GMT+8","ddMMyy"))
      }
      }
}

function cheetahFirstTopUpUsingExcel() {
  var newsheet = SpreadsheetApp.create('充值模板_' + Utilities.formatDate(new Date(), "GMT+8", "ddMMyy")).getActiveSheet()
  //create a new excel
      newsheet.appendRow(header)
      //add the header in
      for (; u < cheetahmessage.length; u++) {
      newsheet.appendRow(cheetahmessage[u])
      }
      //add all the top up / withdrawal requests that belong to cheetah to the empty excel
      var newexcel = newsheet.getParent()
      var url = "https://docs.google.com/feeds/download/spreadsheets/Export?key=" + newexcel.getId() + "&exportFormat=xlsx";
      var params = {
        method      : "get",
        headers     : {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
        muteHttpExceptions: true
      };
      // above is making the exporting the excel 
       var blob1 = UrlFetchApp.fetch(url,params).getBlob()
       newsheet.setName('充值模板_' + Utilities.formatDate(new Date(), "GMT+8", "ddMMyy"))
       blob1.setName(newsheet.getName() + ".xlsx");
       
       //above is renaming the excel and attaching it to the email
       var emailsub = 'adxpert-Facebook-猎豹移动-账户管理-' + Utilities.formatDate(new Date(), "GMT+8","ddMMyy")
       MailApp.sendEmail(cheetahemailaddresses, emailsub, "燕星，麻烦转款/充值/扣减", {attachments: blob1})
      for (var q = 0; q < sheet.getLastRow()-1; q++) {
      if (statuscells[q] == '' && idcells[q] != '' && moneycells[q] != '' && decisioncells[q] != '' && resellercells[q] == 'cheetah') {
      sheet.getRange(q+2, 1).setValue('top up sent to cheetah on ' + Utilities.formatDate(new Date(), "GMT+8","ddMMyy"))
      }
      }
}

function cheetahFirstTopUpEmail() { //roughly the same as yinolink and panda
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
      var emailsub = 'adxpert-Facebook-账户管理-' + Utilities.formatDate(new Date(), "GMT+8","ddMMyy")
      MailApp.sendEmail(cheetahemailaddresses, emailsub, lesserthan2)
      for (var q = 0; q < sheet.getLastRow()-1; q++) {
      if (statuscells[q] == '' && idcells[q] != '' && moneycells[q] != '' && decisioncells[q] != '' && resellercells[q] == 'cheetah') {
      sheet.getRange(q+2, 1).setValue('top up sent to cheetah on ' + Utilities.formatDate(new Date(), "GMT+8","ddMMyy"))
      }
      }
}

function cheetahReplyExcel() { //same as sending the excel the first time, except of sending the email, we are replying the email
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
       
       
       var gmailquary = GmailApp.search('账户管理 ' + Utilities.formatDate(new Date(), "GMT+8", "ddMMyy") + '猎豹移动')
       gmailquary[0].replyAll("燕星，麻烦转款/充值/扣减", {attachments: blob1})
       
       //above are the lines of code that replies the email
      for (var q = 0; q < sheet.getLastRow()-1; q++) {
      if (statuscells[q] == '' && idcells[q] != '' && moneycells[q] != '' && decisioncells[q] != '' && resellercells[q] == 'cheetah') {
      sheet.getRange(q+2, 1).setValue('top up sent to cheetah on ' + Utilities.formatDate(new Date(), "GMT+8","ddMMyy"))
      }
      }
}

function cheetahReplyEmail() { //same as replying the email with the excel sheet, except we are just replying with an email, not sending email
  
      var decisionarray = []
      var lesserthan2 = []
      for (; u < cheetahmessage.length; u++) {
        decisionarray.push(cheetahmessage[u][2])
      
      
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
