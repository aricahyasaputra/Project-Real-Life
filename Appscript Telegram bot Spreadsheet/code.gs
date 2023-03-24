
// How to connect your Telegram Bot to a Google Spreadsheet (Google Apps Script)
// https://www.youtube.com/watch?v=mKSXd_od4Lg
// 
// This code must be added to the Google Apps Script file attached to the spreadsheet script editor. 
// Full steps in the readme

var token = "5508818713:AAFY3fEd8yTNkioqDmGeZSg3i0YZFCjr4vY";     // 1. FILL IN YOUR OWN TOKEN
var telegramUrl = "https://api.telegram.org/bot" + token;
var webAppUrl = "https://script.google.com/macros/s/AKfycbwXSdyaz8iAmehOiVELa19UjQyA7maqLzjYOjgM561g675AMXE/exec"; // 2. FILL IN YOUR GOOGLE WEB APP ADDRESS
var ssId = "1sjlHzYXJuzcxhg3TBM6uPbglGYibdIpzCEIOvigX5qc";      // 3. FILL IN THE ID OF YOUR SPREADSHEET
var adminID = "1325105587";   // 4. Fill in your own Telegram ID for debugging
var sheetName = "App";
var sheet;

function getMe() {
  var url = telegramUrl + "/getMe";
  var response = UrlFetchApp.fetch(url);
  Logger.log(response.getContentText());
}

function setWebhook() {
  var url = telegramUrl + "/setWebhook?url=" + webAppUrl;
  var response = UrlFetchApp.fetch(url);
  Logger.log(response.getContentText());
}

function sendText(id,text) {
  var url = telegramUrl + "/sendMessage?chat_id=" + id + "&text=" + encodeURIComponent(text);
  var response = UrlFetchApp.fetch(url);
  Logger.log(response.getContentText());
}

function doGet(e) {
  return HtmlService.createHtmlOutput("Hi there");
}

function todayReport(){
  var participant = parseInt(SpreadsheetApp.openById(ssId).getSheetByName(sheetName).getRange("AK5").getValues())+5;
  sheet = SpreadsheetApp.openById(ssId).getSheetByName(sheetName).getRange("AB2000:AG2000").getValues();
  
  let jepara = sheet[0][0];
  let pati = sheet[0][1];
  let kudus = sheet[0][2];
  let blora = sheet[0][3];
  let purwodadi = sheet[0][4];
  let total = sheet[0][5];
  
  for(var i = 6;i<participant;i++){
    targetID = SpreadsheetApp.openById(ssId).getSheetByName(sheetName).getRange("AH"+i.toString()).getValues();
    sendText(targetID,"-----Laporan Bulan Ini-----"+"\n- Jepara : "+jepara+"\n- Pati : "+pati+"\n- Kudus : "+kudus+"\n- Blora : "+blora+"\n- Purwodadi : "+purwodadi + "\n Total Witel kudus : " + total); 
  }
}

function doPost(e) {
  try {
    // this is where telegram works
    var data = JSON.parse(e.postData.contents);
    var text = data.message.text;
    var id = data.message.chat.id;
    var name = data.message.chat.first_name + " " + data.message.chat.last_name;
    var answer = "Hi " + name;
    
    
    //sample();
    sendText(id,answer + ",Ketik /help untuk bantuan");
    
    if(text[0]==="/") {
      if(text === "/help"){
        sendText(id,"- /help untuk bantuan\n- /daily untuk menampilkan rata-rata tiap kota hari ini\n- /monthly untuk menampilkan rata-rata tiap kota bulan ini");
      }else if(text === "/daily"){
        sheet = SpreadsheetApp.openById(ssId).getSheetByName(sheetName).getRange("O500:T500").getValues();
        let jepara = sheet[0][0];
        let pati = sheet[0][1];
        let kudus = sheet[0][2];
        let blora = sheet[0][3];
        let purwodadi = sheet[0][4];
        let total = sheet[0][5];
        
        sendText(id,"-----Laporan Hari Ini-----"+"\n- Jepara : "+jepara+"\n- Pati : "+pati+"\n- Kudus : "+kudus+"\n- Blora : "+blora+"\n- Purwodadi : "+purwodadi + "\n Total Witel kudus : " + total); 
      }else if(text === "/monthly"){
        sheet = SpreadsheetApp.openById(ssId).getSheetByName(sheetName).getRange("AB2000:AG2000").getValues();
        let jepara = sheet[0][0];
        let pati = sheet[0][1];
        let kudus = sheet[0][2];
        let blora = sheet[0][3];
        let purwodadi = sheet[0][4];
        let total = sheet[0][5];
        
        sendText(id,"-----Laporan Bulan Ini-----"+"\n- Jepara : "+jepara+"\n- Pati : "+pati+"\n- Kudus : "+kudus+"\n- Blora : "+blora+"\n- Purwodadi : "+purwodadi + "\n Total Witel kudus : " + total); 
      }else if(text.includes('/range')){
        sheet = SpreadsheetApp.openById(ssId).getSheetByName(sheetName).getRange("AB2000:AG2000").getValues();
        let jepara = sheet[0][0];
        let pati = sheet[0][1];
        let kudus = sheet[0][2];
        let blora = sheet[0][3];
        let purwodadi = sheet[0][4];
        let total = sheet[0][5];
        
        sendText(id,"-----Laporan Bulan Ini-----"+"\n- Jepara : "+jepara+"\n- Pati : "+pati+"\n- Kudus : "+kudus+"\n- Blora : "+blora+"\n- Purwodadi : "+purwodadi + "\n Total Witel kudus : " + total); 
      }else{
        sendText(id,"Mohon maaf command yang anda cari tidak ditemukan, gunakan /help untuk bantuan");
      }
     }
    } catch(e) {
    sendText(adminID, JSON.stringify(e,null,4));
  }
}

function addTrigger() {
 ScriptApp.newTrigger("updateCell").timeBased().atHour(20).everyDays(1).create();
}

function updateCell() {
  var cell = "A1";
  SpreadsheetApp.openById(ssId).getRange(cell).setValue("=TO_DATE('EVIDEN DC'!A:A)");
}

function sample(e){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.addEditor('----@-------');
  var name = "/range";
  //var jumlah = parseInt(SpreadsheetApp.openById(ssId).getSheetByName(sheetName).getRange("M6").getValues())+6;
  
  //for(var i = 6;i<jumlah;i++){
    //sheet = SpreadsheetApp.openById(ssId).getSheetByName(sheetName).getRange("I"+i.toString()+":"+"K"+i.toString()).getValues();
    //sendText(adminID,sheet);
  //}str.includes('word')

  if(name.includes('/range')){
    SpreadsheetApp.openById(ssId).getSheetByName(sheetName).getRange("AN3").setValue('hello');
    sendText(adminID,"Berhasil");  
  }else{
    sendText(adminID,"Nope");
  }
}
