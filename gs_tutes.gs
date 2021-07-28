
//Connecting to mongoDB
//POST requests to mongoDB
function postToMongoDb(){
  
  
  var sheet =SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  var range = sheet.getDataRange().getValues();
  
  for(var row =2;row <= sheet.getLastRow();row++){
    //range[row][0]
    var name = sheet.getRange(row, 2).getValue();
    var user_name = sheet.getRange(row, 3).getValue();
    var email = sheet.getRange(row, 4).getValue();
    var phone = sheet.getRange(row, 5).getValue();
    var website = sheet.getRange(row, 6).getValue();
  

    var userData = {
      'name':name,
      'user_name':user_name,
      'email':email,
      'phone_number':phone,
      'website':website
    }
  
    var params = {
      'method':'POST',
      'content-type':'application/json',
      'payload':userData
    }
    
    var url = "https://webhooks.mongodb-realm.com/api/client/v2.0/app/googlesheets-rrtbq/service/google-sheet-connect/incoming_webhook/webhook0";
    var getId = UrlFetchApp.fetch(url,params);
  
    Logger.log(getId);
  }
}


//GET requests to mongoDB

function getToMongoDb(){
  var url = "https://webhooks.mongodb-realm.com/api/client/v2.0/app/googlesheets-rrtbq/service/google-sheet-connect-get/incoming_webhook/webhook0";
  var rawData = UrlFetchApp.fetch(url);
  var jsonData = JSON.parse(rawData);
  var sheet =SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet5");

  for(var row =0;row <= jsonData.length ;row++){
    Logger.log(jsonData[row]._id);
 
    sheet.getRange(row+2, 1).setValue(jsonData[row]._id.$oid);
    sheet.getRange(row+2, 2).setValue(jsonData[row].name);
    sheet.getRange(row+2, 3).setValue(jsonData[row].user_name);
    sheet.getRange(row+2, 4).setValue(jsonData[row].email);
    sheet.getRange(row+2, 5).setValue(jsonData[row].phone_number);
    sheet.getRange(row+2, 6).setValue(jsonData[row].website);
    if(row==0){
    sheet.getRange(1, 1).setValue("Id");
    sheet.getRange(1, 2).setValue("Name");
    sheet.getRange(1, 3).setValue("Username");
    sheet.getRange(1, 4).setValue("Email");
    sheet.getRange(1, 5).setValue("Phone Number");
    sheet.getRange(1, 6).setValue("Website");
    }

  }
}





//Translation using gscripts

//function translate(){
//  var text = "My name is Rock";
//  var translate = LanguageApp.translate(text, "en", "kn");
//  Logger.log(translate);
//}





//Reading texts from the image file

//var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet4");
//var img_column=1;
//var text_column=2;
//
//function getTextFromImage(){
//  var max_row = sheet.getLastRow();
//  
//  for(var row=2;row<=max_row;row++){
//    
//    var url = sheet.getRange(row, img_column).getValue();
//    var imageBlob = UrlFetchApp.fetch(url).getBlob();
//    
//    var resource = {
//      title:imageBlob.getName(),
//      mimeType:imageBlob.getContentType()
//    }
//    
//    var options = {
//      ocr:true
//    }
//    
//    var docFile = Drive.Files.insert(resource,imageBlob,options);
//    var doc = DocumentApp.openById(docFile.id);
//    var text = doc.getBody().getText();
//    
//    sheet.getRange(row, text_column).setValue(text);
//    Drive.Files.remove(docFile.id);
//    
//  }
//   
//  
//}


//Converting googlesheet to JSON
//
//var page_url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTOcuGvruuS95wAT3BJ-obI5BczmoOmmktrUZkOqGYqYmS96yFAVZuesyrAL4TVGBad5yvDc3tHEHvS/pubhtml?gid=0&single=true";
//
//Getting all data from Sheet1
//var json_url = "https://sheetdb.io/api/v1/fgx7rv8wj06ep";
//
//Getting all data from Sheet2
//var json_url_sheet2 = "https://sheetdb.io/api/v1/fgx7rv8wj06ep?sheet=Sheet2";
//
//Searching the contents in database
//var search_sheet2 = "https://sheetdb.io/api/v1/fgx7rv8wj06ep/search?sheet=Sheet2&id=38";
//
//Searching the contents in database
//var search_sheet2 = "https://sheetdb.io/api/v1/fgx7rv8wj06ep/search_or?sheet=Sheet2&userId=2&limit=1";


//Email marketing tool
//
//var url = "https://api.sendinblue.com/v3/smtp/email"; 
//var api_key = "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";
//var sender_name = "CA";
//var sender_email = "caarts.tech@gmail.com";
//var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
//var last_row = sheet.getLastRow();
//var name_col = 1;
//var recipient_col = 2;
//var subject_col = 3;
//var body_col = 4;
//var recipient_email,recipient_name;
//var subject;
//var body;
//var row;
//
//function sendAllEmail(){
//  
//  var headers ={
//    'accept': 'application/json', 
//    'api-key':api_key, 
//    'content-type': 'application/json', 
//    'muteHttpExceptions':true
//  }
//    
//  for(row = 2;row <= last_row;row++){
//    recipient_name = sheet.getRange(row, name_col).getValue(); 
//    recipient_email = sheet.getRange(row, recipient_col).getValue();
//    subject = sheet.getRange(row, subject_col).getValue();
//    body = sheet.getRange(row, body_col).getValue();
//   
//   
//    //Contents to send
//    var data={  
//      "sender":{  
//        "name":sender_name,
//        "email":sender_email
//      },
//      "to":[  
//        {  
//          "email":recipient_email,
//          "name":recipient_name
//        }
//      ],
//      "subject":subject,
//      "htmlContent":"<html><head></head><body><h3>Hi"+recipient_name+" <p>Dear,</p>This is the welcome email sent from CA-Arts.</p></body></html>"
//    }
//   
//   
//    var params = {
//      'method':'POST',
//      'headers':headers,
//      'payload':JSON.stringify(data)
//    }
//    
//    var sendEmail = UrlFetchApp.fetch(url,params);
//    Logger.log(sendEmail.getContentText());
//   
// }
//}






//API Requests and Response
//
//function APIHandler2(){
//  
//  var url ="https://jsonplaceholder.typicode.com/albums";
//  var response = UrlFetchApp.fetch(url);
//  var jsonData = response.getContentText();
//  var data = JSON.parse(jsonData);
//  
//  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet2");
//
//  sheet.getRange(1, 1).setValue("userId");
//  sheet.getRange(1, 2).setValue("id");
//  sheet.getRange(1, 3).setValue("title");
//
//  
//  for(var i=0;i<data.length;i++)
//  {
//    sheet.getRange(i+2, 1).setValue(data[i].userId);
//    sheet.getRange(i+2, 2).setValue(data[i].id);
//    sheet.getRange(i+2, 3).setValue(data[i].title);
//    
//  }
//}














////Email validator
//
//var wb = SpreadsheetApp.getActiveSpreadsheet();
//var sheet1 = wb.getSheetByName('Sheet1');
//
//var apiKey = 'db002a8043d827aa97af545075993d0c';
//var ui = SpreadsheetApp.getUi();
//var email_column = 1;
//var did_you_mean = 2;
//var user = 3;
//var domain = 4;
//var format_valid = 5;
//var mx_found = 6;
//var smtp_check = 7;
//var free = 8;
//var score =9;
//
//
//
//function emailValidator(){
//  
//  for(var row=2;row<=sheet1.getLastRow();row++){
//    var email = sheet1.getRange(row, email_column);
//    var url = "http://apilayer.net/api/check?access_key="+apiKey+"&email="+email+"&smtp=1&format=1";
//    var fetch = UrlFetchApp.fetch(url);
//    
//    if(fetch.getResponseCode()==200){
//      var response = JSON.parse(fetch.getContentText());
//      sheet1.getRange(row, did_you_mean).setValue(response.did_you_mean);
//      sheet1.getRange(row, user).setValue(response.user);
//      sheet1.getRange(row, domain).setValue(response.domain);
//      sheet1.getRange(row, format_valid).setValue(response.format_valid);
//      sheet1.getRange(row, mx_found).setValue(response.mx_found);
//      sheet1.getRange(row, smtp_check).setValue(response.smtp_check);
//      sheet1.getRange(row, free).setValue(response.free);
//      sheet1.getRange(row, score).setValue(response.score);
//    }else{
//      
//    }
//      
//  }
//}



//API Requests and Response
//
//function APIHandler(){
//  
//  var url ="https://jsonplaceholder.typicode.com/users";
//  var response = UrlFetchApp.fetch(url);
//  var jsonData = response.getContentText();
//  var data = JSON.parse(jsonData);
//  
//  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
//
//  sheet.getRange(1, 1).setValue("id");
//  sheet.getRange(1, 2).setValue("name");
//  sheet.getRange(1, 3).setValue("username");
//  sheet.getRange(1, 4).setValue("email");
//  sheet.getRange(1, 5).setValue("phone");
//  sheet.getRange(1, 6).setValue("website");
//  
//  for(var i=0;i<data.length;i++)
//  {
//    sheet.getRange(i+2, 1).setValue(data[i].id);
//    sheet.getRange(i+2, 2).setValue(data[i].name);
//    sheet.getRange(i+2, 3).setValue(data[i].username);
//    sheet.getRange(i+2, 4).setValue(data[i].email);
//    sheet.getRange(i+2, 5).setValue(data[i].phone);
//    sheet.getRange(i+2, 6).setValue(data[i].website);
//    
//  }
//}



//Web Scraping

//function webScraper(){
//  var query="Wiki";
//  var searchResults =UrlFetchApp.fetch("https://www.google.com/search?q="+encodeURIComponent(query));
//  var titleExp = /<a>([\s\S]*?)<\/a>/gi;
//  
//  var results = searchResults.getContentText().match(titleExp);
//  Logger.log(results);
//  
//  
//  
//}


//Sending emails

//function sendEmailToAll() {
//  
//  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
//  var last_row = sheet.getLastRow();
//  var recipient_col = 1;
//  var subject_col = 2;
//  var body_col = 3;
//  var recipient;
//  var subject;
//  var body;
//  var row;
//  
// 
//  for(row = 2;row <= last_row;row++){
//    recipient = sheet.getRange(row, recipient_col).getValue();
//    subject = sheet.getRange(row, subject_col).getValue();
//    body = sheet.getRange(row, body_col).getValue();
//    MailApp.sendEmail(recipient, subject, body);
//    
//  }
//  
//}


//Arrays

//var array1 = ["Rock" ,"Joker","Narayan"];


//Reading all files and folders in Google Drive
//function getAllFilesAndFolders(){
//  
//  var folder = DriveApp.getFolders();
//  var files = DriveApp.getFiles();
//  var colFolders = 1;
//  var colFiles = 2;
//
//  var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
//  
//  for(i=2;folder.hasNext();i++){
//    //returns current folder
//    sheet1.getRange(i, colFolders).setValue(folder.next());
//  }
//  for(i=2;files.hasNext();i++){
//    //returns current folder
//    sheet1.getRange(i, colFiles).setValue(files.next());
//  }
//}


//Reading specific files and folders in Google Drive
//function getSpecificFilesAndFolders(){
//  var file  = DriveApp.getFileById("1jU3bSOtaOM_jFdtgr_AmoaNhq0qSmNkHpHCvWs71VMI");
//  Logger.log(file.getName());
//  Logger.log(file.getUrl());
//  
//}

//Creating folder and Files 
//function createFolderAndFile(){
//  DiveApp.createFolder("Folder1");
//  DriveApp.createFile("Folder1");
//                    
//}


//Search Folders and Files
//function searchFolderAndFiles(){
//  var folders = DriveApp.searchFolders("title contains 'Engineering'");
//  var files = DriveApp.searchFiles("title contains 'img'");
//  
//  if(folders != null){
//    while(folders.hasNext()){
//      Logger.log(folders.next());
//    }
//  }
//  
//  if(files != null){
//    while(files.hasNext()){
//      Logger.log(files.next());
//    }
//  }
//}


////global variable
//var c = 10;
//
//
////Adding Side-bar
//function sideBar(){
//  var html = HtmlService.createHtmlOutputFromFile("test").setTitle("Side-Bar");
//  var ui = SpreadsheetApp.getUi();
//  ui.showSidebar(html);
//  
//}
//
//
////Adding menu
//function onOpen(){
//  
//  var ui = SpreadsheetApp.getUi();
//  
//  ui.createMenu("Extra Menu")
//  .addItem("Side-Bar", "sideBar")
//  .addItem("Add","add").addSeparator()
//  .addSubMenu(ui.createMenu("Extra Sub-Menu").addItem("Extra Menu1","sub1").addItem("Extra Menu2","sub2").addItem("Extra Menu3","sub3"))
//  .addItem("Toastr","tostPrint")
//  .addToUi();
//           
//}
//
//
//function check(){
////if and else
//  var ui = SpreadsheetApp.getUi();
//  if(c>2){
//    
//    ui.alert("Number is greater than 2");
//  }
//  else{
//    ui.alert("Number is less than 2");
//  }
//}
//
////Toast
//function tostPrint(){
//  SpreadsheetApp.getActive().toast("Testing 123...","Test",7);//-1 for permanent
//}
//
////Addition
//function add(){
//  var a,b,res,ch;
//  var response;
//
//  var ui = SpreadsheetApp.getUi();
//  response = ui.prompt("First Number", "Enter first number",ui.ButtonSet.OK_CANCEL);
//  a = response.getResponseText();
//  response = ui.prompt("First Number", "Enter second number",ui.ButtonSet.OK_CANCEL);
//  b=response.getResponseText();
//  res = a+b;
//  ui.alert("Sum :"+res);
//  
//}
//
////Alert
//function sub1() {
//  //local variable
//  var message = "Sub1";
//  SpreadsheetApp.getUi().alert(message);
//}
//
//function sub2() {
//  //local variable
//  var message = "Sub3";
//  SpreadsheetApp.getUi().alert(message);
//}
//
//function sub3() {
//  //local variable
//  var message = "Sub3";
//  SpreadsheetApp.getUi().alert(message);
//}
