function Sendemails() {
  //get a spreadsheet and choose Record of Account Opening Forms, the sheet we are working on
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName("Record of Account Opening Forms"));
  var sheet = SpreadsheetApp.getActiveSheet();
  //choose range from 2nd row and 3rd column and put it into a 2 dimentional array
  var dataRange = sheet.getRange(2,3,sheet.getLastRow()-1,sheet.getLastColumn()-3); 
  //getlastcolumn is only counting the last column that is filled, not from all the way at the 'z' column
  //need put getlastcolumn-3 because infront we starting from the third account
  var data = dataRange.getValues();
  
  var emailsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('email_addresses')
  
  var todaysdate = Utilities.formatDate(new Date(), "GMT+8", "ddMMyy")
  // THE FOLLOWING CODE SHOULD BE CORRECT.
  for (var i = 0;i<sheet.getLastRow()-1;i++) { //if this code fails, replace "sheet.getLastRow()-1" with 998
    //for each (var i in s
    var datecell = sheet.getRange(i+2,3).getValues()
    var statuscell = sheet.getRange(i+2,2).getValues()
    var cheetahorpanda = sheet.getRange(i+2, 6).getValues()
    //Logger.log(idcell)
    if (datecell != "" && statuscell == ""){
      var tableData = data[i]
           
      var excelsheetlabel = sheet.getRange(1, 3, 1, 19).getValues()[0] //get the labels on the excel sheet
      var newsheet = SpreadsheetApp.create('adXpert 广告账号开户 ' + tableData[2]).getActiveSheet().appendRow(
        excelsheetlabel).appendRow(data[i])
      
      
      var newexcel = newsheet.getParent()
      var url = "https://docs.google.com/feeds/download/spreadsheets/Export?key=" + newexcel.getId() + "&exportFormat=xlsx";
      var params = {
        method      : "get",
        headers     : {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
        muteHttpExceptions: true
      };
      var blob1 = UrlFetchApp.fetch(url,params).getBlob()
      newsheet.setName('adXpert 开户表格')
      blob1.setName(newsheet.getName() + ".xlsx");
      
      var idcell2 = sheet.getRange(i+2, 22).getValues().toString()
      
      //do a if statement. for length more than 1, need to split them up then replace. for length = 1, no need split, jus replace. then get blob and attach to email
      if (idcell2.length>1) {
        var textfinderofcomma = idcell2.split(', ')
        var mama2 = []
        var x = 0
        for (var u = 0; u < textfinderofcomma.length; u++) {
          mama2[x] = DriveApp.getFileById(textfinderofcomma[u].replace('https://drive.google.com/open?id=', '')).getBlob()
          x++
        }
        
      }
      else {
        mama2 = DriveApp.getFileById(idcell2.replace('https://drive.google.com/open?id=', '')).getBlob()
      }
      
      var lastblob = []
      lastblob.push(blob1)
      for each (var businessreg in mama2) {
        lastblob.push(businessreg)
      }
      var emailaddress1 = emailsheet.getRange(2, 1).getValues()
      var emailaddress2 = emailsheet.getRange(2,2).getValues()
      var emailaddress3 = emailsheet.getRange(2,3).getValues()
      var emailaddress4 = emailsheet.getRange(2,4).getValues()
      var emailaddress5 = emailsheet.getRange(2,5).getValues()
      var emailaddress6 = emailsheet.getRange(2,6).getValues()
      var emailaddress7 = emailsheet.getRange(2,7).getValues()
      var tempemailadd =  emailaddress4 + ',' + emailaddress5 + ',' + emailaddress6 + ',' + emailaddress7
      var tempemailsub = 'adXpert-Facebook-开账户-' + tableData[7] + '-' + todaysdate
      if (cheetahorpanda == '电商' || cheetahorpanda == '其他' || cheetahorpanda == '科技') {
        
        var realemailaddress = emailaddress1 + ',' + emailaddress2 + ',' + emailaddress3
        var emailsubject = 'adXpert 广告账号开户 ' + tableData[2]
      }
      else if (cheetahorpanda == '游戏或APP') {
        
        var realemailaddress = emailaddress4 + ',' + emailaddress5 + ',' + emailaddress6 + ',' + emailaddress7
        var emailsubject = 'adXpert-Facebook-开账户-' + tableData[7] + '-' + todaysdate
      }
      MailApp.sendEmail(realemailaddress, tempemailsub, "麻烦开户。", {attachments: lastblob});
      
      
      
      sheet.getRange(i + 2, 2).setValue('EMAIL_SENT')
      
      
    };
  }  
} 
   
  


function chineseformreplymailforBMtying() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName("Send emails for BM tying Chinese form"));
  var sheet = SpreadsheetApp.getActiveSheet();
  
  for (var i = 0; i<sheet.getLastRow()-1; i++) {
    var statusbar = sheet.getRange(i+2,1);
    var statusbarvalue = statusbar.getValues()[0][0];
    var ourwhichbm = sheet.getRange(i+2,2).getValues()[0][0];
    var englishnamebar = sheet.getRange(i+2, 4);
    var englishnamevalues = englishnamebar.getValues()[0][0];
    var chinesenamebar = sheet.getRange(i+2,5);
    var chinesenamevalues = chinesenamebar.getValues()[0][0];
    var amountofaccountcreated = Number(sheet.getRange(i+2,6).getValues()[0][0])
    var urlname = sheet.getRange(i+2, 7).getValues()[0][0]
 
      
    //"^[1-9][0-9]{14-15}" /^[1-9][0-9]{14,15}$/
    //searchzhanghaoid is in array format, means if there is more than 1, need to extract them out with some code
    if (statusbarvalue == "" && englishnamevalues !== "") {
      if (urlname.search(",") !== -1) {
      SpreadsheetApp.getUi().alert("There are more than 1 main website. Please remove unnecessary websites so that the acccounts can be named properly.");
      break//i copied this one from the api thing, apparently break is better practice than return, cos break is just break the for loop, and 
      //below function can work
      }
      if (urlname.search("https://www.") !== -1) {
        var httpsindex = urlname.search("https://www.") + 12
        var intermediatiaryname = urlname.slice(httpsindex)
      }
      if (urlname.search("https://") !== -1) {
        var httpsindex = urlname.search("https://") + 8
        var intermediatiaryname = urlname.slice(httpsindex)
      }
      if (urlname.search("http://www.") !== -1) {
        var httpindex = urlname.search("http://www.") + 11
        var intermediatiaryname = urlname.slice(httpindex)
      }
      if (urlname.search("http://") !== -1) {
        var httpindex = urlname.search("http://") + 7
        var intermediatiaryname = urlname.slice(httpindex)
      }
      if (urlname.search("www.") !== -1) {
        var wwwindex = urlname.search("www.") + 4 //plus 4 due to 4 indexes in www.
        var intermediatiaryname = urlname.slice(wwwindex)
      }
      if (urlname.search("www.") == -1 || urlname.search("http://") == -1 || urlname.search("http://www.") == -1 || urlname.search("https://") == -1 || urlname.search("https://www.") == -1) {
      var withoutdotcomindex = urlname.search(/\./)
      var newaccountname = urlname.slice(0, withoutdotcomindex)
      }
      if (intermediatiaryname) {
      var secondFullStopIndex = intermediatiaryname.search(/\./)
      var newaccountname = intermediatiaryname.slice(0, secondFullStopIndex)
      }
      var gmailsheet = GmailApp.search(englishnamevalues)[0]; //[0] ensures the latest email thread, 0 is always counting from top
      if (gmailsheet == undefined) {
        var gmailsheet2 = GmailApp.search(chinesenamevalues)[0]; //all in all, its counting from the top.
        var gmailsheet2body = gmailsheet2.getMessages()[gmailsheet2.getMessageCount()-1].getBody() //getmessagecount a number which allows us to get the latest message
        //then it needs to minus 1 because the messages from the top starts from 0, if get the message count will be 1 more than the maximum index in the count, so need minus 1
        var allnumbersthatis15or16digits = gmailsheet2body.match(/[0-9]{15,16}/g) //.toString() // single slash / is a special notation for regular expression
        var gmailsheet2BRid = []
        for (k = 0; k < amountofaccountcreated; k++) {
          gmailsheet2BRid.push(allnumbersthatis15or16digits[k])
          }
        gmailsheet2BRid.toString() 
        if (ourwhichbm == 'a') {
          gmailsheet2.replyAll(gmailsheet2BRid + "\n\n" + "请绑定" + sheet.getRange(1, 10).getValues()[0][0] + "和" + sheet.getRange(i+2,3).getValues()[0][0] + "\n\n" + "麻烦改名为" + newaccountname)
        }
        if (ourwhichbm == 'g') {
          gmailsheet2.replyAll(gmailsheet2BRid + "\n\n" + "请绑定" + sheet.getRange(2, 10).getValues()[0][0] + "和" + sheet.getRange(i+2,3).getValues()[0][0] + "\n\n" + "麻烦改名为" + newaccountname)
        }
        if (ourwhichbm == 'o') {
          gmailsheet2.replyAll(gmailsheet2BRid + "\n\n" + "请绑定" + sheet.getRange(3, 10).getValues()[0][0] + "和" + sheet.getRange(i+2,3).getValues()[0][0]+ "\n\n" + "麻烦改名为" + newaccountname)
        }
        if (ourwhichbm == 'b') {
          gmailsheet2.replyAll(gmailsheet2BRid + "\n\n" + "请绑定" + sheet.getRange(4, 10).getValues()[0][0] + "和" + sheet.getRange(i+2,3).getValues()[0][0]+ "\n\n" + "麻烦改名为" + newaccountname)
        }
        if (ourwhichbm == 'i') {
          gmailsheet2.replyAll(gmailsheet2BRid + "\n\n" + "请绑定" + sheet.getRange(5, 10).getValues()[0][0] + "和" + sheet.getRange(i+2,3).getValues()[0][0]+ "\n\n" + "麻烦改名为" + newaccountname)
        }
        statusbar.setValue("BMs tied")
      }
      else {
        var gmailsheetbody =  gmailsheet.getMessages()[gmailsheet.getMessageCount()-1].getBody()
        var allnumbersthatis15or16digits = gmailsheetbody.match(/[0-9]{15,16}/g)//.toString() //need to b able to ignore other numbrs and print out the acct numbers only
        var gmailsheetBRid = []
        for (k = 0; k < amountofaccountcreated; k++) {
          
          gmailsheetBRid.push(allnumbersthatis15or16digits[k])
                 
         }
        gmailsheetBRid.toString() 
        if (ourwhichbm == 'a') {
           gmailsheet.replyAll(gmailsheetBRid + "\n\n" + "请绑定" + sheet.getRange(1, 10).getValues()[0][0] + "和" + sheet.getRange(i+2,3).getValues()[0][0]+ "\n\n" + "麻烦改名为" + newaccountname)
         }
        if (ourwhichbm == 'g') {
          gmailsheet.replyAll(gmailsheetBRid + "\n\n" + "请绑定" + sheet.getRange(2, 10).getValues()[0][0] + "和" + sheet.getRange(i+2,3).getValues()[0][0]+ "\n\n" + "麻烦改名为" + newaccountname)
        }
        if (ourwhichbm == 'o') {
          gmailsheet.replyAll(gmailsheetBRid + "\n\n" + "请绑定" + sheet.getRange(3, 10).getValues()[0][0] + "和" + sheet.getRange(i+2,3).getValues()[0][0]+ "\n\n" + "麻烦改名为" + newaccountname)
        }
        if (ourwhichbm == 'b') {
          gmailsheet.replyAll(gmailsheetBRid + "\n\n" + "请绑定" + sheet.getRange(4, 10).getValues()[0][0] + "和" + sheet.getRange(i+2,3).getValues()[0][0]+ "\n\n" + "麻烦改名为" + newaccountname)
        }
        if (ourwhichbm == 'i') {
          gmailsheet.replyAll(gmailsheetBRid + "\n\n" + "请绑定" + sheet.getRange(5, 10).getValues()[0][0] + "和" + sheet.getRange(i+2,3).getValues()[0][0]+ "\n\n" + "麻烦改名为" + newaccountname)
        }
        statusbar.setValue("BMs tied")
      }
    }
  }
}
  
      
        
      

