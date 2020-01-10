/* Facebook Reporting & Cost Data Upload in Google Analytics 
 * Description: Exports Facebook Ads Data in Google Sheets & Uploads it To Google Analytics.
 * @Ritwikga www.Digishuffle.com
 *
 * Updated: 21-05-2019
 * - Breakdown Feature
 * - Data Import Alerts
 * - UI Improvements & Bug Fixes
 *
 * Recent Updates @https://github.com/RitwikGA/FacebookReportingTool/
 */

///// Facebook Details ///////

var CLIENT_ID = '';     // Insert App ID                                        

var CLIENT_SECRET = '';   // Insert App Secret

//var FB_AD_ACCOUNT_ID = '';   //Ad Account Id

// More fields at https://developers.facebook.com/docs/marketing-api/insights/parameters 
var FB_FIELDS = 'name,account_id,account_status,amount_spent,balance,business,disable_reason,spend_cap'; 



var pos = [1,1]     //Spreadsheet Cell Position





/**
*
*  Input Variable Values Ends
*
*/


////// ACCOUNTDATA Literal ////////////////////
var ACCOUNTDATA = {
adAccountUIFields : ['name','account_status','amount_spent','balance','business','disable_reason','spend_cap'],
   /// The Columns To Be Populated in the Fields Box in the UI.
 
getUIFields : function(y) {return y.map(function(i){
                                        return {id:i,text:i.split('_').map(function(j){
                                        return j.charAt(0).toUpperCase()+j.slice(1)}).join(' ')}})},
getUIHeaders : function(k){return k.map(function(i){
return i.split('_').map(function(j){
return j.charAt(0).toUpperCase()+j.slice(1)}).join(' ')})},
facebookData : {//facebookAccountId:FB_AD_ACCOUNT_ID,
                
                facebookFields:FB_FIELDS},
                //facebookBreakdowns:FB_BREAKDOWN},

//dateData : {
  //preDefinedRage: DATE_RANGE
  
//},                

additionalData: {


pos:pos,
//limit:limit
}
}

function showBar() {
var html=HtmlService.createTemplateFromFile('digiSideBar').evaluate().setTitle("Facebook Reporting Tool").setWidth(300)
SpreadsheetApp.getUi().showSidebar(html)
}

function facebookData()
{ makeRequest(ACCOUNTDATA)
Logger.log('Current project has ' + ScriptApp.getProjectTriggers().length + ' triggers.');

}
  


function onOpen() {
SpreadsheetApp.getUi().createMenu('Reports').addSubMenu(SpreadsheetApp.getUi()
 .createMenu('Facebook').addItem("Open Sidebar", 'showBar').addSeparator().addItem("Authorize", 'fbAuth').addItem("Log Out", 'reset').addItem("Export Data", 'facebookData').addItem("Account Balancing", 'creditCalculation'))
.addSeparator().addItem("oAuth Redirect URI", 'getValidOauthRedirectUrl').addToUi();

}

function fbAuth(){
  var UI=HtmlService.createTemplate("<b><a href='<?=getService().getAuthorizationUrl()?>' target='_blank'>Click To Authorize</a></b><br /><? if(getService().hasAccess())"+ 
                                    "{ ?> <?!= <p><span style='color:green'>Authorized Successfully</span></p> } else {?> <?!= <p><span style='color:red'>Not Authorized</span></p> }").evaluate() 
  SpreadsheetApp.getUi().showModalDialog(UI, "Facebook Authorization")
 
}


function jsonToQuery(param) //"key"is an important word, so the tool designer wants to allocate key = something, and the something is related to the parameters in param.
{
var str = "";
for (var key in param) {
    if (str != "") {
        str += "&";
    }
    str += key + "=" + param[key];
}
return str
}

function getValidOauthRedirectUrl(){ //this function is just getting permission to allow data to be used. 

var validOauthUrl = Utilities.formatString('https://developers.facebook.com/apps/%s/fb-login/settings/', CLIENT_ID)
var htmlOutput = HtmlService
.createHtmlOutput('<style>span{font-size: 14px;font-weight: bold;text-decoration: underline;font-style: italic;cursor: pointer;}</style>'+
                  '<script>function selectURL(){document.getElementById("oauthURL").select();' + 
                  'document.execCommand("copy")}</script>'+
                  '<p>Copy & Paste The Below URL In <a href="'+validOauthUrl+'">' + 
                  '<i>Valid OAuth Redirect URIs</i></a></p>'+
                  '<br /><span onclick="selectURL()">Copy to Clipboard</span><textarea type="text" id="oauthURL" style="width:100%;">'+Utilities.formatString("https://script.google.com/macros/d/%s/usercallback", ScriptApp.getScriptId())+
                  '</textarea>') //html code that allows the validoauthurl to be shown in the text box, with all the fonts, and character size
    .setWidth(450)
    .setHeight(200);
SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Valid OAuth Redirect URIs')
}


function makeRequest(ACCOUNTDATAOBJECT) {

//if(!ACCOUNTDATAOBJECT['callFrom']) {ACCOUNTDATAOBJECT = ACCOUNTDATA}  

 var fbRequest = getService(); 
 var requestEndpoint = "https://graph.facebook.com/v4.0/me/adaccounts/?fields=insights.date_preset(lifetime),account_id,spend_cap,balance,amount_spent,name,account_status,business,disable_reason" //most probably must change from /insights? to nothing
 //var param = {'limit':ACCOUNTDATAOBJECT['additionalData']['limit']}//'level': ACCOUNTDATAOBJECT['facebookData']['facebookLevel']}
 
 //****** PROBABLY NEED TO CHANGE THIS ENTIRE PART HERE ABOUT PARAM!!! -**************/////////
 

 var param= {'fields': ACCOUNTDATAOBJECT['facebookData']['facebookFields']}
 

if(param['fields'] == ''){SpreadsheetApp.getUi().alert("Enter The Fields");return}
  



//********** AT LEAST UNTIL HERE!!!*************//
 
var accToken = "EAAFZAa43WTj4BAAJRZARD8Q059UbMLAfcaVaq6LzlAmY9mbu2SwG7XJKLnzgnNpX7VY9u1Bq7gokwvdkghanOneVdaPrryZAb7rFn0HmHFuVLVcSVtVSWeuicjpZCacnCnn7JuS2Rz67MnGcSRNR7vhtuELpNHVP022Rt3rCzAZDZD";
var response = UrlFetchApp.fetch(requestEndpoint, 
{headers: {'Authorization': 'Bearer ' + accToken},muteHttpExceptions : true})
// response above is just getting the data from the facebook ad account
// response will be in string
var parseData = JSON.parse(response) //as the response is in string, need to json parse it to make it into a json file
if(parseData.hasOwnProperty('error'))
{
if(parseData.error.hasOwnProperty('error_user_title'))
{SpreadsheetApp.getUi().alert(parseData.error.error_user_title)}
else{SpreadsheetApp.getUi().alert(parseData.error.message)}
return
}

//if(parseData.data.length == 0)
//{SpreadsheetApp.getUi().alert('No Facebook Data For The Applied Date Range'); return;}  
  


 

try{
parseData = nextTokenData(parseData)
var fieldArray = param['fields'].split(",")
//if(param['breakdowns']){ = fieldArray.concat(param['breakdowns'].split(","))}

var headers = ACCOUNTDATA.getUIHeaders(fieldArray)
headers.push('Maximum Credit');
headers.push('Amount Left in Account Now')
var sheet= SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LIVE');  


var pos = ACCOUNTDATAOBJECT['additionalData']['pos']  
if(typeof(ACCOUNTDATAOBJECT['additionalData']['pos']) == 'string'){
pos = ACCOUNTDATAOBJECT['additionalData']['pos'].split(",")   
}  

if(sheet.getLastRow() > 0 && sheet.getLastColumn() > 0)
{sheet.getRange(pos[0],pos[1],sheet.getLastRow(),sheet.getLastColumn()).clear()}

var finalParsedOutput = []    


finalParsedOutput = parserNonGA(parseData, fieldArray);
finalParsedOutput.unshift(headers);
  
var cell = sheet.getRange(pos[0],pos[1],1,finalParsedOutput[0].length)  
cell.setFontWeight("bold")
cell.setBorder(false, false, true, false, false, false,"black",SpreadsheetApp.BorderStyle.DOUBLE)

  
if(finalParsedOutput[0].length > 0) {sheet.getRange(pos[0], pos[1], finalParsedOutput.length, finalParsedOutput[0].length).setValues(finalParsedOutput)}

var statusDescription = "ACCOUNT_ID: "+ACCOUNTDATAOBJECT['facebookData']['facebookAccountId']+"<br /> ROWS: "+finalParsedOutput.length;
return {status:'success', description:statusDescription}
} catch (e) {Logger.log(e) }
  

};
 

function nextTokenData(parseData)
{
var accToken = "EAAFZAa43WTj4BAAJRZARD8Q059UbMLAfcaVaq6LzlAmY9mbu2SwG7XJKLnzgnNpX7VY9u1Bq7gokwvdkghanOneVdaPrryZAb7rFn0HmHFuVLVcSVtVSWeuicjpZCacnCnn7JuS2Rz67MnGcSRNR7vhtuELpNHVP022Rt3rCzAZDZD"
var datalength123 = Object.keys(parseData).length
if(datalength123 == 0) {return parseData}    
var fbRequest = getService();
if(parseData.paging.next != undefined)
{ 
var parsedata_pg = parseData;
while (true)
{
var response = UrlFetchApp.fetch(parsedata_pg.paging.next, 
  {headers: {'Authorization': 'Bearer ' + accToken},muteHttpExceptions : true})
parsedata_pg = JSON.parse(response)
parseData.data = parseData.data.concat(parsedata_pg.data)
if(parsedata_pg.paging.next == undefined) 
{ break;}
}}
return parseData
}
//as the json have lots of pages (only specific for facebook), this one is about putting all the data from pages into one page


function parserNonGA(parseData, fieldsArray){
//var data=parseData.data;  
var rw=[];
//var test1 = parseData[2] //tested and object cannot be identify by number
//for (var i = 0; i < Object.keys(parseData).length; i++)
//{
  fieldsArray.push("max_credit");
  fieldsArray.push("amount_left")
  parseData.data.map(function(adact, index){
    adact.spend_cap /= 100;
    adact.amount_spent /= 100;
    if (adact.insights) {
        adact.balance = parseFloat(adact.insights.data[0].spend);
    }
    adact.max_credit = parseFloat(adact.spend_cap) - parseFloat(adact.amount_spent) + parseFloat(adact.balance);
    adact.amount_left = parseFloat(adact.spend_cap) - parseFloat(adact.amount_spent)
    rw[index]= Array.apply(null, new Array(fieldsArray.length)).map(Number.prototype.valueOf,0);

                for (key in adact) {
    var mama = key
    rw[index][fieldsArray.indexOf(key)] = adact[key]}//.replace(/\,|\'|\"/g,'')} //basically all the parseData should be with parseData[i]
});
  

//}
return rw
}

function deleteTimeEventTriggers() {
  //var newsheet = SpreadsheetApp.getActiveSpreadsheet()
  //.insertSheet('marketing api drawn on ' + Utilities.formatDate(new Date(), "GMT+8", "ddMMyy")) //dont wanna create a new spreadsheet, but just wanna create a new sheet
  //var ss = newsheet.getParent()
  var triggers = ScriptApp.getProjectTriggers();
for (var i = 0; i < triggers.length; i++) {
  ScriptApp.deleteTrigger(triggers[i]);
}
  
    //.forSpreadsheet(SpreadsheetApp.create('marketing api drawn on ' + Utilities.formatDate(new Date(), "GMT+8", "ddMMyy"))).onOpen().create()
 } 
 
function createTimeEventTriggers(){ScriptApp.newTrigger('facebookData').timeBased().everyMinutes(1).create();}
  
function duplicateLiveSheet(){
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  ss.setActiveSheet(ss.getSheetByName('LIVE'))
  var newsheet = SpreadsheetApp.getActiveSpreadsheet().duplicateActiveSheet();
  newsheet.setName(Utilities.formatDate(new Date(), "GMT+8", "ddMMyy"));
}

function createTriggerForNewSheet(){
  ScriptApp.newTrigger('duplicateLiveSheet').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).create();
}


function creditCalculation () {
  //const adAccountID = ''
  //const facebookAppID = '379793452715582'
  const facebookAppSecret = '81500d1d938af11bd92be8633d30c3f5'
  const facebookFields = 'name,account_id,account_status,amount_spent,balance,spend_cap'
  const position = [1,2]
  const accountData = {//adAccountUIFields : ['name','account_status','amount_spent','balance','business','spend_cap'],
   /// The Columns To Be Populated in the Fields Box in the UI.
 
getUIFields : function(y) {return y.map(function(i){return {id:i,text:i.split('_').map(function(j){return j.charAt(0).toUpperCase()+j.slice(1)}).join(' ')}})},
getUIHeaders : function(k){return k.map(function(i){return i.split('_').map(function(j){return j.charAt(0).toUpperCase()+j.slice(1)}).join(' ')})},
facebookData : {//facebookAccountId:facebookAppID,
                facebookFields:facebookFields},
                additionalData: {pos:position,} 
}
  const facebookRequest = getService();
  const requestEndpoint = "https://graph.facebook.com/v4.0/me/adaccounts/?fields=insights.date_preset(lifetime),account_id,spend_cap,balance,amount_spent,name,account_status,business,disable_reason"
  const param= {'fields': accountData['facebookData']['facebookFields']}
  const accToken = "EAAFZAa43WTj4BAAJRZARD8Q059UbMLAfcaVaq6LzlAmY9mbu2SwG7XJKLnzgnNpX7VY9u1Bq7gokwvdkghanOneVdaPrryZAb7rFn0HmHFuVLVcSVtVSWeuicjpZCacnCnn7JuS2Rz67MnGcSRNR7vhtuELpNHVP022Rt3rCzAZDZD";
  const response = UrlFetchApp.fetch(requestEndpoint, 
  {headers: {'Authorization': 'Bearer ' + accToken},muteHttpExceptions : true})
  var parseData = JSON.parse(response) //as the response is in string, need to json parse it to make it into a json file
  if(parseData.hasOwnProperty('error')) {
    if(parseData.error.hasOwnProperty('error_user_title'))
      {SpreadsheetApp.getUi().alert(parseData.error.error_user_title)}
    else{SpreadsheetApp.getUi().alert(parseData.error.message)}
    return
}
  //try{
  const dataLength = Object.keys(parseData).length
  if(dataLength == 0) {return parseData}  
  var parsedataNextPage = parseData.paging.next
  //if(parsedataNextPage != undefined)
   //{ 
  //var parsedata_pg = parseData;
  while (parsedataNextPage !== undefined)
  {
  //var parsedata_pg = parseData;
  var response2 = UrlFetchApp.fetch(parsedataNextPage, 
    {headers: {'Authorization': 'Bearer ' + accToken},muteHttpExceptions : true})
  var parsedata_pgs = JSON.parse(response2)
   parseData.data = parseData.data.concat(parsedata_pgs.data)
  //if(parsedata_pg.paging.next == undefined) 
  //{ break;}
  parsedataNextPage = parsedata_pgs.paging.next
  }
  if (parseData.paging.next == undefined) {
  return parseData
  }
  const fieldArray = param['fields'].split(",")
  const headers = fieldArray
   //headers.push('Maximum Credit');
  //headers.push('Amount Left in Account Now')
  const sheet= SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Ad Acc Bank + Data');  


  
 

if(sheet.getLastRow() > 0 && sheet.getLastColumn() > 0)
{sheet.getRange(position[0],position[1],sheet.getLastRow(),sheet.getLastColumn()).clear()}
  
  var finalParsedOutput = []   
  
  //var rw=[];
//var test1 = parseData[2] //tested and object cannot be identify by number
//for (var i = 0; i < Object.keys(parseData).length; i++)
//{
  fieldArray.push("max_credit");
  fieldArray.push("amount_left")
  parseData.data.map(function(adact, index){
    adact.spend_cap /= 100;
    adact.amount_spent /= 100;
    if (adact.insights) {
        adact.balance = parseFloat(adact.insights.data[0].spend);
    }
    adact.max_credit = parseFloat(adact.spend_cap) - parseFloat(adact.amount_spent) + parseFloat(adact.balance);
    adact.amount_left = parseFloat(adact.spend_cap) - parseFloat(adact.amount_spent)
    finalParsedOutput[index]= Array.apply(null, new Array(fieldArray.length)).map(Number.prototype.valueOf,0);

                for (key in adact) {
    var mama = key
    finalParsedOutput[index][fieldArray.indexOf(key)] = adact[key]}//.replace(/\,|\'|\"/g,'')} //basically all the parseData should be with parseData[i]
});
  



finalParsedOutput.unshift(headers);
  
var cell = sheet.getRange(position[0],position[1],1,finalParsedOutput[0].length)  
cell.setFontWeight("bold")
cell.setBorder(false, false, true, false, false, false,"black",SpreadsheetApp.BorderStyle.DOUBLE)

  
if(finalParsedOutput[0].length > 0) {sheet.getRange(position[0], position[1], finalParsedOutput.length, finalParsedOutput[0].length).setValues(finalParsedOutput)}

//var statusDescription = "ACCOUNT_ID: "+ACCOUNTDATAOBJECT['facebookData']['facebookAccountId']+"<br /> ROWS: "+finalParsedOutput.length;
//return {status:'success', description:statusDescription}
 //catch (e) {Logger.log(e) }
}

/**
*  oAuth Script : https://github.com/googlesamples/apps-script-oauth2
*/

/**
 * Configures the service.
 */
function getService() {
  return OAuth2.createService('Facebook')
      // Set the endpoint URLs.
      .setAuthorizationBaseUrl('https://www.facebook.com/dialog/oauth')
      .setTokenUrl('https://graph.facebook.com/v4.0/oauth/access_token')

      // Set the client ID and secret.
      .setClientId(CLIENT_ID)
      .setClientSecret(CLIENT_SECRET)

      // Set the name of the callback function that should be invoked to complete
      // the OAuth flow.
      .setCallbackFunction('authCallback')
  
  
      //Set Scope
      .setScope('ads_read') 
  

      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getUserProperties());
}


function authCallback(request) {
  var isAuthorized = getService().handleCallback(request);
 
  if (isAuthorized) {
    successUI(true)
    showBar()
    return HtmlService.createHtmlOutput('Success! You can close this tab.<script>window.top.close()</script>');
  } else {
    successUI(false)
    showBar()
    return HtmlService.createHtmlOutput('Denied. You can close this tab.<script>window.top.close()</script>');
  }
}


function reset() {
  var service = getService();
  service.reset();
  showBar()
  SpreadsheetApp.getUi().alert("Log Out Success!!")
}

function successUI(isAuth){

 if(isAuth){
  var UI=HtmlService.createHtmlOutput("<b><span style='color:green'>Authorization Successful</span></b>")
  SpreadsheetApp.getUi().showModalDialog(UI, "Authorization Status") } else
  {var UI=HtmlService.createHtmlOutput("<b><span style='color:red'>Authorization Fail</span></b>")
    SpreadsheetApp.getUi().showModalDialog(UI, "Authorization Status")}
}


function adAccounts(){

 var fbRequest = getService(); 
  var addaccounts_endpoint = "https://graph.facebook.com/v4.0/me?fields=adaccounts.limit(100)%7Bname,account_id%7D" 
 
  var adAccountInfo = UrlFetchApp.fetch(addaccounts_endpoint, 
  {headers: {'Authorization': 'Bearer ' + fbRequest.getAccessToken()},muteHttpExceptions : true}) 
  
  var parsedadAccountInfo = JSON.parse(adAccountInfo)
  if(parsedadAccountInfo.hasOwnProperty('error') || !parsedadAccountInfo.adaccounts)
  {SpreadsheetApp.getUi().alert('ERROR: '+parsedadAccountInfo['error']['message']);return false} 
  else {
  var adAccountFB = nextTokenData(parsedadAccountInfo.adaccounts,200)
  var parsed_adurls = parsedadAccountInfo;
  parsed_adurls['adaccounts'] = adAccountFB
}


  return { 'facebookAccountData':parsed_adurls.adaccounts.data } 
}


////
//
//Cost Data Upload Script - http://www.ryanpraski.com/google-analytics-cost-data-import-google-sheets-automated/ 
//
////






//////////////////////////////////////////////
(function (host, expose) {
   var module = { exports: {} };
   var exports = module.exports;
   /****** code begin *********/


/**
 * @file Contains the methods exposed by the library, and performs
 * any required setup.
 */

/**
 * The supported formats for the returned OAuth2 token.
 * @enum {string}
 */
var TOKEN_FORMAT = {
  /** JSON format, for example <code>{"access_token": "..."}</code> **/
  JSON: 'application/json',
  /** Form URL-encoded, for example <code>access_token=...</code> **/
  FORM_URL_ENCODED: 'application/x-www-form-urlencoded'
};

/**
 * The supported locations for passing the state parameter.
 * @enum {string}
 */
var STATE_PARAMETER_LOCATION = {
  /**
   * Pass the state parameter in the authorization URL.
   * @default
   */
  AUTHORIZATION_URL: 'authorization-url',
  /**
   * Pass the state token in the redirect URL, as a workaround for APIs that
   * don't support the state parameter.
   */
  REDIRECT_URL: 'redirect-url'
};

/**
 * Creates a new OAuth2 service with the name specified. It's usually best to
 * create and configure your service once at the start of your script, and then
 * reference them during the different phases of the authorization flow.
 * @param {string} serviceName The name of the service.
 * @return {Service_} The service object.
 */
function createService(serviceName) {
  return new Service_(serviceName);
}

/**
 * Returns the redirect URI that will be used for a given script. Often this URI
 * needs to be entered into a configuration screen of your OAuth provider.
 * @param {string} scriptId The script ID of your script, which can be found in
 *     the Script Editor UI under "File > Project properties".
 * @return {string} The redirect URI.
 */
function getRedirectUri(scriptId) {
  return Utilities.formatString(
    'https://script.google.com/macros/d/%s/usercallback', scriptId);
}

if (typeof module === 'object') {
  module.exports = {
    createService: createService,
    getRedirectUri: getRedirectUri,
    TOKEN_FORMAT: TOKEN_FORMAT,
    STATE_PARAMETER_LOCATION: STATE_PARAMETER_LOCATION
  };
}


var Service_ = function(serviceName) {
  validate_({
    'Service name': serviceName
  });
  this.serviceName_ = serviceName;
  this.params_ = {};
  this.tokenFormat_ = TOKEN_FORMAT.JSON;
  this.tokenHeaders_ = null;
  this.scriptId_ = eval('Script' + 'App').getScriptId();
  this.expirationMinutes_ = 60;
};

/**
 * The number of seconds before a token actually expires to consider it expired
 * and refresh it.
 * @type {number}
 * @private
 */
Service_.EXPIRATION_BUFFER_SECONDS_ = 60;

/**
 * The number of milliseconds that a token should remain in the cache.
 * @type {number}
 * @private
 */
Service_.LOCK_EXPIRATION_MILLISECONDS_ = 30 * 1000;

/**
 * Sets the service's authorization base URL (required). For Google services
 * this URL should be
 * https://accounts.google.com/o/oauth2/auth.
 * @param {string} authorizationBaseUrl The authorization endpoint base URL.
 * @return {Service_} This service, for chaining.
 */
Service_.prototype.setAuthorizationBaseUrl = function(authorizationBaseUrl) {
  this.authorizationBaseUrl_ = authorizationBaseUrl;
  return this;
};

/**
 * Sets the service's token URL (required). For Google services this URL should
 * be https://accounts.google.com/o/oauth2/token.
 * @param {string} tokenUrl The token endpoint URL.
 * @return {Service_} This service, for chaining.
 */
Service_.prototype.setTokenUrl = function(tokenUrl) {
  this.tokenUrl_ = tokenUrl;
  return this;
};

/**
 * Sets the service's refresh URL. Some OAuth providers require a different URL
 * to be used when generating access tokens from a refresh token.
 * @param {string} refreshUrl The refresh endpoint URL.
 * @return {Service_} This service, for chaining.
 */
Service_.prototype.setRefreshUrl = function(refreshUrl) {
  this.refreshUrl_ = refreshUrl;
  return this;
};

/**
 * Sets the format of the returned token. Default: OAuth2.TOKEN_FORMAT.JSON.
 * @param {OAuth2.TOKEN_FORMAT} tokenFormat The format of the returned token.
 * @return {Service_} This service, for chaining.
 */
Service_.prototype.setTokenFormat = function(tokenFormat) {
  this.tokenFormat_ = tokenFormat;
  return this;
};

/**
 * Sets the additional HTTP headers that should be sent when retrieving or
 * refreshing the access token.
 * @param {Object.<string,string>} tokenHeaders A map of header names to values.
 * @return {Service_} This service, for chaining.
 */
Service_.prototype.setTokenHeaders = function(tokenHeaders) {
  this.tokenHeaders_ = tokenHeaders;
  return this;
};

/**
 * @callback tokenHandler
 * @param tokenPayload {Object} A hash of parameters to be sent to the token
 *     URL.
 * @param tokenPayload.code {string} The authorization code.
 * @param tokenPayload.client_id {string} The client ID.
 * @param tokenPayload.client_secret {string} The client secret.
 * @param tokenPayload.redirect_uri {string} The redirect URI.
 * @param tokenPayload.grant_type {string} The type of grant requested.
 * @returns {Object} A modified hash of parameters to be sent to the token URL.
 */

/**
 * Sets an additional function to invoke on the payload of the access token
 * request.
 * @param {tokenHandler} tokenHandler tokenHandler A function to invoke on the
 *     payload of the request for an access token.
 * @return {Service_} This service, for chaining.
 */
Service_.prototype.setTokenPayloadHandler = function(tokenHandler) {
  this.tokenPayloadHandler_ = tokenHandler;
  return this;
};

/**
 * Sets the name of the authorization callback function (required). This is the
 * function that will be called when the user completes the authorization flow
 * on the service provider's website. The callback accepts a request parameter,
 * which should be passed to this service's <code>handleCallback()</code> method
 * to complete the process.
 * @param {string} callbackFunctionName The name of the callback function.
 * @return {Service_} This service, for chaining.
 */
Service_.prototype.setCallbackFunction = function(callbackFunctionName) {
  this.callbackFunctionName_ = callbackFunctionName;
  return this;
};

/**
 * Sets the client ID to use for the OAuth flow (required). You can create
 * client IDs in the "Credentials" section of a Google Developers Console
 * project. Although you can use any project with this library, it may be
 * convinient to use the project that was created for your script. These
 * projects are not visible if you visit the console directly, but you can
 * access it by click on the menu item "Resources > Advanced Google services" in
 * the Script Editor, and then click on the link "Google Developers Console" in
 * the resulting dialog.
 * @param {string} clientId The client ID to use for the OAuth flow.
 * @return {Service_} This service, for chaining.
 */
Service_.prototype.setClientId = function(clientId) {
  this.clientId_ = clientId;
  return this;
};

/**
 * Sets the client secret to use for the OAuth flow (required). See the
 * documentation for <code>setClientId()</code> for more information on how to
 * create client IDs and secrets.
 * @param {string} clientSecret The client secret to use for the OAuth flow.
 * @return {Service_} This service, for chaining.
 */
Service_.prototype.setClientSecret = function(clientSecret) {
  this.clientSecret_ = clientSecret;
  return this;
};

/**
 * Sets the property store to use when persisting credentials (required). In
 * most cases this should be user properties, but document or script properties
 * may be appropriate if you want to share access across users.
 * @param {PropertiesService.Properties} propertyStore The property store to use
 *     when persisting credentials.
 * @return {Service_} This service, for chaining.
 * @see https://developers.google.com/apps-script/reference/properties/
 */
Service_.prototype.setPropertyStore = function(propertyStore) {
  this.propertyStore_ = propertyStore;
  return this;
};

/**
 * Sets the cache to use when persisting credentials (optional). Using a cache
 * will reduce the need to read from the property store and may increase
 * performance. In most cases this should be a private cache, but a public cache
 * may be appropriate if you want to share access across users.
 * @param {CacheService.Cache} cache The cache to use when persisting
 *     credentials.
 * @return {Service_} This service, for chaining.
 * @see https://developers.google.com/apps-script/reference/cache/
 */
Service_.prototype.setCache = function(cache) {
  this.cache_ = cache;
  return this;
};

/**
 * Sets the lock to use when checking and refreshing credentials (optional).
 * Using a lock will ensure that only one execution will be able to access the
 * stored credentials at a time. This can prevent race conditions that arise
 * when two executions attempt to refresh an expired token.
 * @param {LockService.Lock} lock The lock to use when accessing credentials.
 * @return {Service_} This service, for chaining.
 * @see https://developers.google.com/apps-script/reference/lock/
 */
Service_.prototype.setLock = function(lock) {
  this.lock_ = lock;
  return this;
};

/**
 * Sets the scope or scopes to request during the authorization flow (optional).
 * If the scope value is an array it will be joined using the separator before
 * being sent to the server, which is is a space character by default.
 * @param {string|Array.<string>} scope The scope or scopes to request.
 * @param {string} [optSeparator] The optional separator to use when joining
 *     multiple scopes. Default: space.
 * @return {Service_} This service, for chaining.
 */
Service_.prototype.setScope = function(scope, optSeparator) {
  var separator = optSeparator || ' ';
  this.params_.scope = Array.isArray(scope) ? scope.join(separator) : scope;
  return this;
};

/**
 * Sets an additional parameter to use when constructing the authorization URL
 * (optional). See the documentation for your service provider for information
 * on what parameter values they support.
 * @param {string} name The parameter name.
 * @param {string} value The parameter value.
 * @return {Service_} This service, for chaining.
 */
Service_.prototype.setParam = function(name, value) {
  this.params_[name] = value;
  return this;
};

/**
 * Sets the private key to use for Service Account authorization.
 * @param {string} privateKey The private key.
 * @return {Service_} This service, for chaining.
 */
Service_.prototype.setPrivateKey = function(privateKey) {
  this.privateKey_ = privateKey;
  return this;
};

/**
 * Sets the issuer (iss) value to use for Service Account authorization.
 * If not set the client ID will be used instead.
 * @param {string} issuer This issuer value
 * @return {Service_} This service, for chaining.
 */
Service_.prototype.setIssuer = function(issuer) {
  this.issuer_ = issuer;
  return this;
};

/**
 * Sets the subject (sub) value to use for Service Account authorization.
 * @param {string} subject This subject value
 * @return {Service_} This service, for chaining.
 */
Service_.prototype.setSubject = function(subject) {
  this.subject_ = subject;
  return this;
};

/**
 * Sets number of minutes that a token obtained through Service Account
 * authorization should be valid. Default: 60 minutes.
 * @param {string} expirationMinutes The expiration duration in minutes.
 * @return {Service_} This service, for chaining.
 */
Service_.prototype.setExpirationMinutes = function(expirationMinutes) {
  this.expirationMinutes_ = expirationMinutes;
  return this;
};

/**
 * Gets the authorization URL. The first step in getting an OAuth2 token is to
 * have the user visit this URL and approve the authorization request. The
 * user will then be redirected back to your application using callback function
 * name specified, so that the flow may continue.
 * @return {string} The authorization URL.
 */
Service_.prototype.getAuthorizationUrl = function() {
  validate_({
    'Client ID': this.clientId_,
    'Script ID': this.scriptId_,
    'Callback function name': this.callbackFunctionName_,
    'Authorization base URL': this.authorizationBaseUrl_
  });

  var redirectUri = getRedirectUri(this.scriptId_);
  var state = eval('Script' + 'App').newStateToken()
      .withMethod(this.callbackFunctionName_)
      .withArgument('serviceName', this.serviceName_)
      .withTimeout(3600)
      .createToken();
  var params = {
    client_id: this.clientId_,
    response_type: 'code',
    redirect_uri: redirectUri,
    state: state
  };
  params = extend_(params, this.params_);
  return buildUrl_(this.authorizationBaseUrl_, params);
};

/**
 * Completes the OAuth2 flow using the request data passed in to the callback
 * function.
 * @param {Object} callbackRequest The request data recieved from the callback
 *     function.
 * @return {boolean} True if authorization was granted, false if it was denied.
 */
Service_.prototype.handleCallback = function(callbackRequest) {
  var code = callbackRequest.parameter.code;
  var error = callbackRequest.parameter.error;
  if (error) {
    if (error == 'access_denied') {
      return false;
    } else {
      throw new Error('Error authorizing token: ' + error);
    }
  }
  validate_({
    'Client ID': this.clientId_,
    'Client Secret': this.clientSecret_,
    'Script ID': this.scriptId_,
    'Token URL': this.tokenUrl_
  });
  var redirectUri = getRedirectUri(this.scriptId_);
  var headers = {
    'Accept': this.tokenFormat_
  };
  if (this.tokenHeaders_) {
    headers = extend_(headers, this.tokenHeaders_);
  }
  var tokenPayload = {
    code: code,
    client_id: this.clientId_,
    client_secret: this.clientSecret_,
    redirect_uri: redirectUri,
    grant_type: 'authorization_code'
  };
  if (this.tokenPayloadHandler_) {
    tokenPayload = this.tokenPayloadHandler_(tokenPayload);
  }
  var response = UrlFetchApp.fetch(this.tokenUrl_, {
    method: 'post',
    headers: headers,
    payload: tokenPayload,
    muteHttpExceptions: true
  });
  var token = this.getTokenFromResponse_(response);
  this.saveToken_(token);
  return true;
};

/**
 * Determines if the service has access (has been authorized and hasn't
 * expired). If offline access was granted and the previous token has expired
 * this method attempts to generate a new token.
 * @return {boolean} true if the user has access to the service, false
 *     otherwise.
 */
Service_.prototype.hasAccess = function() {
  return this.lockable_(function() {
    var token = this.getToken();
    if (!token || this.isExpired_(token)) {
      if (token && token.refresh_token) {
        try {
          this.refresh();
        } catch (e) {
          this.lastError_ = e;
          return false;
        }
      } else if (this.privateKey_) {
        try {
          this.exchangeJwt_();
        } catch (e) {
          this.lastError_ = e;
          return false;
        }
      } else {
        return false;
      }
    }
    return true;
  });
};

/**
 * Gets an access token for this service. This token can be used in HTTP
 * requests to the service's endpoint. This method will throw an error if the
 * user's access was not granted or has expired.
 * @return {string} An access token.
 */
Service_.prototype.getAccessToken = function() {
  if (!this.hasAccess()) {
    throw new Error('Access not granted or expired.');
  }
  var token = this.getToken();
  return token.access_token;
};

/**
 * Resets the service, removing access and requiring the service to be
 * re-authorized.
 */
Service_.prototype.reset = function() {
  this.getStorage().removeValue(null);
};

/**
 * Gets the last error that occurred this execution when trying to automatically
 * refresh or generate an access token.
 * @return {Exception} An error, if any.
 */
Service_.prototype.getLastError = function() {
  return this.lastError_;
};

/**
 * Returns the redirect URI that will be used for this service. Often this URI
 * needs to be entered into a configuration screen of your OAuth provider.
 * @return {string} The redirect URI.
 */
Service_.prototype.getRedirectUri = function() {
  return getRedirectUri(this.scriptId_);
};

/**
 * Gets the token from a UrlFetchApp response.
 * @param {UrlFetchApp.HTTPResponse} response The response object.
 * @return {Object} The parsed token.
 * @throws If the token cannot be parsed or the response contained an error.
 * @private
 */
Service_.prototype.getTokenFromResponse_ = function(response) {
  var token = this.parseToken_(response.getContentText());
  var resCode = response.getResponseCode();
  if ( resCode < 200 || resCode >= 300 || token.error) {
    var reason = [
      token.error,
      token.message,
      token.error_description,
      token.error_uri
    ].filter(Boolean).map(function(part) {
      return typeof(part) == 'string' ? part : JSON.stringify(part);
    }).join(', ');
    if (!reason) {
      reason = resCode + ': ' + JSON.stringify(token);
    }
    throw new Error('Error retrieving token: ' + reason);
  }
  return token;
};

/**
 * Parses the token using the service's token format.
 * @param {string} content The serialized token content.
 * @return {Object} The parsed token.
 * @private
 */
Service_.prototype.parseToken_ = function(content) {
  var token;
  if (this.tokenFormat_ == TOKEN_FORMAT.JSON) {
    try {
      token = JSON.parse(content);
    } catch (e) {
      throw new Error('Token response not valid JSON: ' + e);
    }
  } else if (this.tokenFormat_ == TOKEN_FORMAT.FORM_URL_ENCODED) {
    token = content.split('&').reduce(function(result, pair) {
      var parts = pair.split('=');
      result[decodeURIComponent(parts[0])] = decodeURIComponent(parts[1]);
      return result;
    }, {});
  } else {
    throw new Error('Unknown token format: ' + this.tokenFormat_);
  }
  token.granted_time = getTimeInSeconds_(new Date());
  return token;
};

/**
 * Refreshes a token that has expired. This is only possible if offline access
 * was requested when the token was authorized.
 */
Service_.prototype.refresh = function() {
  validate_({
    'Client ID': this.clientId_,
    'Client Secret': this.clientSecret_,
    'Token URL': this.tokenUrl_
  });

  this.lockable_(function() {
    var token = this.getToken();
    if (!token.refresh_token) {
      throw new Error('Offline access is required.');
    }
    var headers = {
      Accept: this.tokenFormat_
    };
    if (this.tokenHeaders_) {
      headers = extend_(headers, this.tokenHeaders_);
    }
    var tokenPayload = {
        refresh_token: token.refresh_token,
        client_id: this.clientId_,
        client_secret: this.clientSecret_,
        grant_type: 'refresh_token'
    };
    if (this.tokenPayloadHandler_) {
      tokenPayload = this.tokenPayloadHandler_(tokenPayload);
    }
    // Use the refresh URL if specified, otherwise fallback to the token URL.
    var url = this.refreshUrl_ || this.tokenUrl_;
    var response = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: headers,
      payload: tokenPayload,
      muteHttpExceptions: true
    });
    var newToken = this.getTokenFromResponse_(response);
    if (!newToken.refresh_token) {
      newToken.refresh_token = token.refresh_token;
    }
    this.saveToken_(newToken);
  });
};

/**
 * Gets the storage layer for this service, used to persist tokens.
 * Custom values associated with the service can be stored here as well.
 * The key <code>null</code> is used to to store the token and should not
 * be used.
 * @return {Storage} The service's storage.
 */
Service_.prototype.getStorage = function() {
  validate_({
    'Property store': this.propertyStore_
  });
  if (!this.storage_) {
    var prefix = 'oauth2.' + this.serviceName_;
    this.storage_ = new Storage_(prefix, this.propertyStore_, this.cache_);
  }
  return this.storage_;
};

/**
 * Saves a token to the service's property store and cache.
 * @param {Object} token The token to save.
 * @private
 */
Service_.prototype.saveToken_ = function(token) {
  this.getStorage().setValue(null, token);
};

/**
 * Gets the token from the service's property store or cache.
 * @return {Object} The token, or null if no token was found.
 */
Service_.prototype.getToken = function() {
  return this.getStorage().getValue(null);
};

/**
 * Determines if a retrieved token is still valid.
 * @param {Object} token The token to validate.
 * @return {boolean} True if it has expired, false otherwise.
 * @private
 */
Service_.prototype.isExpired_ = function(token) {
  var expiresIn = token.expires_in || token.expires;
  if (!expiresIn) {
    return false;
  } else {
    var expiresTime = token.granted_time + Number(expiresIn);
    var now = getTimeInSeconds_(new Date());
    return expiresTime - now < Service_.EXPIRATION_BUFFER_SECONDS_;
  }
};

/**
 * Uses the service account flow to exchange a signed JSON Web Token (JWT) for
 * an access token.
 * @private
 */
Service_.prototype.exchangeJwt_ = function() {
  validate_({
    'Token URL': this.tokenUrl_
  });
  var jwt = this.createJwt_();
  var headers = {
    'Accept': this.tokenFormat_
  };
  if (this.tokenHeaders_) {
    headers = extend_(headers, this.tokenHeaders_);
  }
  var response = UrlFetchApp.fetch(this.tokenUrl_, {
    method: 'post',
    headers: headers,
    payload: {
      assertion: jwt,
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer'
    },
    muteHttpExceptions: true
  });
  var token = this.getTokenFromResponse_(response);
  this.saveToken_(token);
};

/**
 * Creates a signed JSON Web Token (JWT) for use with Service Account
 * authorization.
 * @return {string} The signed JWT.
 * @private
 */
Service_.prototype.createJwt_ = function() {
  validate_({
    'Private key': this.privateKey_,
    'Token URL': this.tokenUrl_,
    'Issuer or Client ID': this.issuer_ || this.clientId_
  });
  var header = {
    alg: 'RS256',
    typ: 'JWT'
  };
  var now = new Date();
  var expires = new Date(now.getTime());
  expires.setMinutes(expires.getMinutes() + this.expirationMinutes_);
  var claimSet = {
    iss: this.issuer_ || this.clientId_,
    aud: this.tokenUrl_,
    exp: Math.round(expires.getTime() / 1000),
    iat: Math.round(now.getTime() / 1000)
  };
  if (this.subject_) {
    claimSet.sub = this.subject_;
  }
  if (this.params_.scope) {
    claimSet.scope = this.params_.scope;
  }
  var toSign = Utilities.base64EncodeWebSafe(JSON.stringify(header)) + '.' +
      Utilities.base64EncodeWebSafe(JSON.stringify(claimSet));
  var signatureBytes =
      Utilities.computeRsaSha256Signature(toSign, this.privateKey_);
  var signature = Utilities.base64EncodeWebSafe(signatureBytes);
  return toSign + '.' + signature;
};

/**
 * Locks access to a block of code if a lock has been set on this service.
 * @param {function} func The code to execute.
 * @return {*} The result of the code block.
 * @private
 */
Service_.prototype.lockable_ = function(func) {
  var releaseLock = false;
  if (this.lock_ && !this.lock_.hasLock()) {
    this.lock_.waitLock(Service_.LOCK_EXPIRATION_MILLISECONDS_);
    releaseLock = true;
  }
  var result = func.apply(this);
  if (this.lock_ && releaseLock) {
    this.lock_.releaseLock();
  }
  return result;
};

// Copyright 2017 Google Inc. All Rights Reserved.
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

/**
 * @file Contains classes used to persist data and access it.
 */

/**
 * Creates a new Storage_ instance, which is used to persist OAuth tokens and
 * related information.
 * @param {string} prefix The prefix to use for keys in the properties and
 *     cache.
 * @param {PropertiesService.Properties} properties The properties instance to
 *     use.
 * @param {CacheService.Cache} [optCache] The optional cache instance to use.
 * @constructor
 */
function Storage_(prefix, properties, optCache) {
  this.prefix_ = prefix;
  this.properties_ = properties;
  this.cache_ = optCache;
  this.memory_ = {};
}

/**
 * The TTL for cache entries, in seconds.
 * @type {number}
 * @private
 */
Storage_.CACHE_EXPIRATION_TIME_SECONDS = 21600; // 6 hours.

/**
 * Gets a stored value.
 * @param {string} key The key.
 * @return {*} The stored value.
 */
Storage_.prototype.getValue = function(key) {
  // Check memory.
  if (this.memory_[key]) {
    return this.memory_[key];
  }

  var prefixedKey = this.getPrefixedKey_(key);
  var jsonValue;
  var value;

  // Check cache.
  if (this.cache_ && (jsonValue = this.cache_.get(prefixedKey))) {
    value = JSON.parse(jsonValue);
    this.memory_[key] = value;
    return value;
  }

  // Check properties.
  if (jsonValue = this.properties_.getProperty(prefixedKey)) {
    if (this.cache_) {
      this.cache_.put(prefixedKey,
          jsonValue, Storage_.CACHE_EXPIRATION_TIME_SECONDS);
    }
    value = JSON.parse(jsonValue);
    this.memory_[key] = value;
    return value;
  }

  // Not found.
  return null;
};

/**
 * Stores a value.
 * @param {string} key The key.
 * @param {*} value The value.
 */
Storage_.prototype.setValue = function(key, value) {
  var prefixedKey = this.getPrefixedKey_(key);
  var jsonValue = JSON.stringify(value);
  this.properties_.setProperty(prefixedKey, jsonValue);
  if (this.cache_) {
    this.cache_.put(prefixedKey, jsonValue,
        Storage_.CACHE_EXPIRATION_TIME_SECONDS);
  }
  this.memory_[key] = value;
};

/**
 * Removes a stored value.
 * @param {string} key The key.
 */
Storage_.prototype.removeValue = function(key) {
  var prefixedKey = this.getPrefixedKey_(key);
  this.properties_.deleteProperty(prefixedKey);
  if (this.cache_) {
    this.cache_.remove(prefixedKey);
  }
  delete this.memory_[key];
};

/**
 * Gets a key with the prefix applied.
 * @param {string} key The key.
 * @return {string} The key with the prefix applied.
 * @private
 */
Storage_.prototype.getPrefixedKey_ = function(key) {
  if (key) {
    return this.prefix_ + '.' + key;
  } else {
    return this.prefix_;
  }
};
// Copyright 2014 Google Inc. All Rights Reserved.
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

/**
 * @file Contains utility methods used by the library.
 */

/* exported buildUrl_ */
/**
 * Builds a complete URL from a base URL and a map of URL parameters.
 * @param {string} url The base URL.
 * @param {Object.<string, string>} params The URL parameters and values.
 * @return {string} The complete URL.
 * @private
 */
function buildUrl_(url, params) {
  var paramString = Object.keys(params).map(function(key) {
    return encodeURIComponent(key) + '=' + encodeURIComponent(params[key]);
  }).join('&');
  return url + (url.indexOf('?') >= 0 ? '&' : '?') + paramString;
}

/* exported validate_ */
/**
 * Validates that all of the values in the object are non-empty. If an empty
 * value is found, and error is thrown using the key as the name.
 * @param {Object.<string, string>} params The values to validate.
 * @private
 */
function validate_(params) {
  Object.keys(params).forEach(function(name) {
    var value = params[name];
    if (!value) {
      throw Utilities.formatString('%s is required.', name);
    }
  });
}

/* exported getTimeInSeconds_ */
/**
 * Gets the time in seconds, rounded down to the nearest second.
 * @param {Date} date The Date object to convert.
 * @return {Number} The number of seconds since the epoch.
 * @private
 */
function getTimeInSeconds_(date) {
  return Math.floor(date.getTime() / 1000);
}

/* exported extend_ */
/**
 * Copy all of the properties in the source objects over to the
 * destination object, and return the destination object.
 * @param {Object} destination The combined object.
 * @param {Object} source The object who's properties are copied to the
 *     destination.
 * @return {Object} A combined object with the desination and source
 *     properties.
 * @see http://underscorejs.org/#extend
 */
function extend_(destination, source) {
  var keys = Object.keys(source);
  for (var i = 0; i < keys.length; ++i) {
    destination[keys[i]] = source[keys[i]];
  }
  return destination;
}

   /****** code end *********/
   ;(
function copy(src, target, obj) {
    obj[target] = obj[target] || {};
    if (src && typeof src === 'object') {
        for (var k in src) {
            if (src.hasOwnProperty(k)) {
                obj[target][k] = src[k];
            }
        }
    } else {
        obj[target] = src;
    }
}
   ).call(null, module.exports, expose, host);
}).call(this, this, "OAuth2");
 
