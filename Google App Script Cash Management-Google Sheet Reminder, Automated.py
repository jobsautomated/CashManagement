function onEdit() {
  var file = DriveApp.getFilesByName("REPLACEBANK2Balance.csv").next();
  var date= file.getLastUpdated();
  var current= new Date()
  var difference=(current.getTime()-date.getTime())/60000
  var fedline = DriveApp.getFilesByName("FedlineStatus.csv").next();
  var fed= fedline.getLastUpdated();
  var cad = DriveApp.getFileById('1Dx-IW6SCghxgRetyePEsIzKlTaaZ-U_e').getLastUpdated();
  var hourly = Utilities.formatDate(new Date(),"GMT-4:00", "HH:mm");
  var lastdate = Utilities.formatDate(date,"GMT-4:00", "MM/dd/yyyy");
  var caddate = Utilities.formatDate(cad,"GMT-4:00", "MM/dd/yyyy");
  var feddate = Utilities.formatDate(fed,"GMT-4:00", "MM/dd/yyyy");
  var today = Utilities.formatDate(new Date(),"GMT-4:00", "MM/dd/yyyy");
  var todayformat = Utilities.formatDate(new Date(),"GMT-4:00", "M/d/yyyy");
  if(difference>1.006)
  {
    if (hourly =="09:08" & lastdate==today)
  {
    emailme(today);
  }
    else if (hourly =="14:05" & caddate==today)
  {
    convert(today);
  }
    else if (hourly =="17:18")
  {
    getTextFromPDF();
  }
    else if ((hourly =="18:03"||hourly =="18:18"||hourly =="18:33"||hourly =="18:48"||hourly =="19:03"||hourly =="19:18"||hourly =="19:33"||hourly =="19:48"||hourly =="20:03"||hourly =="20:18"||hourly =="20:33"||hourly =="20:48")& feddate==today)
  {
    FedlineStatus(today);
  }
  else if (hourly =="08:48")
  {
    BankRecon();
  }
  else if (hourly =="08:58")
  {
    dashboard();
  }
  else if (hourly =="07:58")
  {
    expectedpayment(todayformat);
  }
    Logger.log(difference);
    return;
  }
  EOD(hourly)
  }

function QuantumMatch(){
 
 var sourceFolderId = "1w6uHIqvHkGI0sid7vkAjpc0S2hBlc1SY"; // Folder ID including source files.
  var destinationFolderId = "1w6uHIqvHkGI0sid7vkAjpc0S2hBlc1SY"; // Folder ID that the converted files are put.

  var getFileIds = function (folder, fileList, q) {
    var files = folder.searchFiles(q);
    while (files.hasNext()) {
      var f = files.next();
      fileList.push({id: f.getId(), fileName: f.getName().split(".")[0].trim()});
    }
    var folders = folder.getFolders();
    while (folders.hasNext()) getFileIds(folders.next(), fileList, q);
    return fileList;
  };
  var sourceFiles = getFileIds(DriveApp.getFolderById(sourceFolderId), [], "mimeType='" + MimeType.MICROSOFT_EXCEL + "' or mimeType='" + MimeType.MICROSOFT_EXCEL_LEGACY + "'");
  var destinationFiles = getFileIds(DriveApp.getFolderById(destinationFolderId), [], "mimeType='" + MimeType.GOOGLE_SHEETS + "'");
  var createFiles = sourceFiles.filter(function(e) {return destinationFiles.every(function(f) {return f.fileName !== e.fileName});});
  var updateFiles = sourceFiles.reduce(function(ar, e) {
    var dst = destinationFiles.filter(function(f) {return f.fileName === e.fileName});
    if (dst.length > 0) {
      e.to = dst[0].id;
      ar.push(e);
    }
    return ar;
  }, []);
  if (createFiles.length > 0) createFiles.forEach(function(e) {Drive.Files.insert({mimeType: MimeType.GOOGLE_SHEETS, parents: [{id: destinationFolderId}], title: e.fileName}, DriveApp.getFileById(e.id))});
  if (updateFiles.length > 0) updateFiles.forEach(function(e) {Drive.Files.update({}, e.to, DriveApp.getFileById(e.id))});
 var sh = SpreadsheetApp.openById(updateFiles[0].to).getSheets()[0];
 var data = sh.getDataRange().getDisplayValues();
 var sheet = SpreadsheetApp.openById('1haf4IYeIfUJL8ClNzPjSxMtTQp32IAKTSl1SXQO0DHo').getSheets()[0];
 sheet.clear()
 sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
 var sh1 = SpreadsheetApp.openById(updateFiles[0].to).getSheets()[1];
 var data1 = sh1.getDataRange().getDisplayValues();
 var sheet1 = SpreadsheetApp.openById('1w7qUXwp4cBxxGncWYrmslEpYpSqjV0z9P7XKlw2u8F4').getSheets()[0];
 sheet1.clear()
 sheet1.getRange(1, 1, data1.length, data1[0].length).setValues(data1);
}

function EOD(hourly) {
  //CSV to Sheet
  var file = DriveApp.getFilesByName("REPLACEBANK1Balance.csv").next();
  var date= file.getLastUpdated()
  var REPLACEBANK1 = Utilities.parseCsv(file.getBlob().getDataAsString());
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.clear()
  sheet.getRange(1, 1, REPLACEBANK1.length, REPLACEBANK1[0].length).setValues(REPLACEBANK1);
  //sheet.insertRowsAfter(REPLACEBANK1.length+1, 1);
  var file1 = DriveApp.getFilesByName("REPLACEBANK2Balance.csv").next();
  var REPLACEBANK2 = Utilities.parseCsv(file1.getBlob().getDataAsString());
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(sheet.getLastRow()+1,1,REPLACEBANK2.length,REPLACEBANK2[0].length).setValues(REPLACEBANK2);
  sendEmail(REPLACEBANK2.length,hourly);
}

function FedlineStatus(today) {
  //CSV to Sheet
  var file = DriveApp.getFilesByName("FedlineStatus.csv").next();
  var Status = Utilities.parseCsv(file.getBlob().getDataAsString());
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(100, 1, Status.length, Status[0].length).setValues(Status);
  var currentstatus= SpreadsheetApp.getActiveSheet().getRange(105,2).getDisplayValue();
  var currentstatusExt= SpreadsheetApp.getActiveSheet().getRange(104,2).getDisplayValue();
  var date= SpreadsheetApp.getActiveSheet().getRange(108,3).getDisplayValue();
  var dateREPLACEBANK2= SpreadsheetApp.getActiveSheet().getRange(1,1).getDisplayValue();
  var msg= SpreadsheetApp.getActiveSheet().getRange(110,1).getDisplayValue();
  //Logger.log(currentstatusExt.split("\"")[1])
  var ss = SpreadsheetApp.openById("1xz1hheHURpwkRojtyfbz02YVbpXmIh9pz_Ns8YfICdg");
  var sheets = ss.getSheets();
  var datestatus = sheets[0].getRange(2,2).getDisplayValue();
  Logger.log(datestatus)
  Logger.log(today)
  //sheet.getRange(lastRow,lastColumn).setValue('Pending');
  if(datestatus!=today){
    if(currentstatusExt.split("\"")[1]=="Service Issue" && Number(date.split(",")[0].split(" ")[1])!=Number(dateREPLACEBANK2.split(",")[0].split(": ")[1].split(" ")[1]))
  {
    extension(currentstatusExt.split("\"")[1]);
  }
    else if (Number(date.split(",")[0].split(" ")[1])==Number(dateREPLACEBANK2.split(",")[0].split(": ")[1].split(" ")[1]))
  {
     emailstatus(currentstatus.split("\"")[1],date,msg,today);
  }
  }  
  return;
}


function sendEmail(REPLACEBANK2,hourly)
{
  var REPLACEBANK2data=REPLACEBANK2+4;
  var date=SpreadsheetApp.getActiveSheet().getRange(1,1).getValue();
  var REPLACEBANK1 = SpreadsheetApp.getActiveSheet().getRange(5,2).getDisplayValue();
  var REPLACEBANK2 = SpreadsheetApp.getActiveSheet().getRange(46,2).getDisplayValue();
  Logger.log(REPLACEBANK2data);
  var Total = SpreadsheetApp.getActiveSheet().getRange(5,2).getValue()+SpreadsheetApp.getActiveSheet().getRange(46,2).getValue();
  //var Total = parseFloat(REPLACEBANK1.split("**")[REPLACEBANK1.split("**").length-1].split("$")[1].split(",").join(""))+parseFloat(REPLACEBANK2.split("**")[REPLACEBANK2.split("**").length-1].split("$")[1].split(",").join(""))
  var Curr = formatCurrency('$', Total);
  if (hourly =="17:01"||hourly =="17:02"||hourly =="17:03"||hourly =="17:04"||hourly =="17:05")
  {
  fivepm(date,REPLACEBANK2,REPLACEBANK1,Curr);
  }
  else if (hourly =="18:31"||hourly =="18:32"||hourly =="18:33")
  {
  sixthirty(date,REPLACEBANK2,REPLACEBANK1,Curr);
  }
  else if (hourly =="99:99")
  {
  sixthirty(date,REPLACEBANK2,REPLACEBANK1,Curr);
  }
  else
  {
  composeApprovedEmail(date,REPLACEBANK2,REPLACEBANK1,Curr);
  }
}

function formatCurrency(symbol, amount) {
  var aDigits = amount.toFixed(2).split(".");
  aDigits[0] = aDigits[0].split("").reverse().join("")
    .replace(/(\d{3})(?=\d)/g,"$1,").split("").reverse().join("");
  return symbol + aDigits.join(".");
}
 
 
function composeApprovedEmail(dt,REPLACEBANK2,REPLACEBANK1,total)
{

  var email = "REPLACE";
  //var email = "function onEdit() {
  var file = DriveApp.getFilesByName("REPLACEBANK2Balance.csv").next();
  var date= file.getLastUpdated();
  var current= new Date()
  var difference=(current.getTime()-date.getTime())/60000
  var fedline = DriveApp.getFilesByName("FedlineStatus.csv").next();
  var fed= fedline.getLastUpdated();
  var cad = DriveApp.getFileById('1Dx-IW6SCghxgRetyePEsIzKlTaaZ-U_e').getLastUpdated();
  var hourly = Utilities.formatDate(new Date(),"GMT-4:00", "HH:mm");
  var lastdate = Utilities.formatDate(date,"GMT-4:00", "MM/dd/yyyy");
  var caddate = Utilities.formatDate(cad,"GMT-4:00", "MM/dd/yyyy");
  var feddate = Utilities.formatDate(fed,"GMT-4:00", "MM/dd/yyyy");
  var today = Utilities.formatDate(new Date(),"GMT-4:00", "MM/dd/yyyy");
  var todayformat = Utilities.formatDate(new Date(),"GMT-4:00", "M/d/yyyy");
  if(difference>1.006)
  {
    if (hourly =="09:08" & lastdate==today)
  {
    emailme(today);
  }
    else if (hourly =="14:05" & caddate==today)
  {
    convert(today);
  }
    else if (hourly =="17:18")
  {
    getTextFromPDF();
  }
    else if ((hourly =="18:03"||hourly =="18:18"||hourly =="18:33"||hourly =="18:48"||hourly =="19:03"||hourly =="19:18"||hourly =="19:33"||hourly =="19:48"||hourly =="20:03"||hourly =="20:18"||hourly =="20:33"||hourly =="20:48")& feddate==today)
  {
    FedlineStatus(today);
  }
  else if (hourly =="08:48")
  {
    BankRecon();
  }
  else if (hourly =="08:58")
  {
    dashboard();
  }
  else if (hourly =="07:58")
  {
    expectedpayment(todayformat);
  }
    Logger.log(difference);
    return;
  }
  EOD(hourly)
  }

function QuantumMatch(){
 
 var sourceFolderId = "1w6uHIqvHkGI0sid7vkAjpc0S2hBlc1SY"; // Folder ID including source files.
  var destinationFolderId = "1w6uHIqvHkGI0sid7vkAjpc0S2hBlc1SY"; // Folder ID that the converted files are put.

  var getFileIds = function (folder, fileList, q) {
    var files = folder.searchFiles(q);
    while (files.hasNext()) {
      var f = files.next();
      fileList.push({id: f.getId(), fileName: f.getName().split(".")[0].trim()});
    }
    var folders = folder.getFolders();
    while (folders.hasNext()) getFileIds(folders.next(), fileList, q);
    return fileList;
  };
  var sourceFiles = getFileIds(DriveApp.getFolderById(sourceFolderId), [], "mimeType='" + MimeType.MICROSOFT_EXCEL + "' or mimeType='" + MimeType.MICROSOFT_EXCEL_LEGACY + "'");
  var destinationFiles = getFileIds(DriveApp.getFolderById(destinationFolderId), [], "mimeType='" + MimeType.GOOGLE_SHEETS + "'");
  var createFiles = sourceFiles.filter(function(e) {return destinationFiles.every(function(f) {return f.fileName !== e.fileName});});
  var updateFiles = sourceFiles.reduce(function(ar, e) {
    var dst = destinationFiles.filter(function(f) {return f.fileName === e.fileName});
    if (dst.length > 0) {
      e.to = dst[0].id;
      ar.push(e);
    }
    return ar;
  }, []);
  if (createFiles.length > 0) createFiles.forEach(function(e) {Drive.Files.insert({mimeType: MimeType.GOOGLE_SHEETS, parents: [{id: destinationFolderId}], title: e.fileName}, DriveApp.getFileById(e.id))});
  if (updateFiles.length > 0) updateFiles.forEach(function(e) {Drive.Files.update({}, e.to, DriveApp.getFileById(e.id))});
 var sh = SpreadsheetApp.openById(updateFiles[0].to).getSheets()[0];
 var data = sh.getDataRange().getDisplayValues();
 var sheet = SpreadsheetApp.openById('1haf4IYeIfUJL8ClNzPjSxMtTQp32IAKTSl1SXQO0DHo').getSheets()[0];
 sheet.clear()
 sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
 var sh1 = SpreadsheetApp.openById(updateFiles[0].to).getSheets()[1];
 var data1 = sh1.getDataRange().getDisplayValues();
 var sheet1 = SpreadsheetApp.openById('1w7qUXwp4cBxxGncWYrmslEpYpSqjV0z9P7XKlw2u8F4').getSheets()[0];
 sheet1.clear()
 sheet1.getRange(1, 1, data1.length, data1[0].length).setValues(data1);
}

function EOD(hourly) {
  //CSV to Sheet
  var file = DriveApp.getFilesByName("REPLACEBANK1Balance.csv").next();
  var date= file.getLastUpdated()
  var REPLACEBANK1 = Utilities.parseCsv(file.getBlob().getDataAsString());
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.clear()
  sheet.getRange(1, 1, REPLACEBANK1.length, REPLACEBANK1[0].length).setValues(REPLACEBANK1);
  //sheet.insertRowsAfter(REPLACEBANK1.length+1, 1);
  var file1 = DriveApp.getFilesByName("REPLACEBANK2Balance.csv").next();
  var REPLACEBANK2 = Utilities.parseCsv(file1.getBlob().getDataAsString());
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(sheet.getLastRow()+1,1,REPLACEBANK2.length,REPLACEBANK2[0].length).setValues(REPLACEBANK2);
  sendEmail(REPLACEBANK2.length,hourly);
}

function FedlineStatus(today) {
  //CSV to Sheet
  var file = DriveApp.getFilesByName("FedlineStatus.csv").next();
  var Status = Utilities.parseCsv(file.getBlob().getDataAsString());
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(100, 1, Status.length, Status[0].length).setValues(Status);
  var currentstatus= SpreadsheetApp.getActiveSheet().getRange(105,2).getDisplayValue();
  var currentstatusExt= SpreadsheetApp.getActiveSheet().getRange(104,2).getDisplayValue();
  var date= SpreadsheetApp.getActiveSheet().getRange(108,3).getDisplayValue();
  var dateREPLACEBANK2= SpreadsheetApp.getActiveSheet().getRange(1,1).getDisplayValue();
  var msg= SpreadsheetApp.getActiveSheet().getRange(110,1).getDisplayValue();
  //Logger.log(currentstatusExt.split("\"")[1])
  var ss = SpreadsheetApp.openById("1xz1hheHURpwkRojtyfbz02YVbpXmIh9pz_Ns8YfICdg");
  var sheets = ss.getSheets();
  var datestatus = sheets[0].getRange(2,2).getDisplayValue();
  Logger.log(datestatus)
  Logger.log(today)
  //sheet.getRange(lastRow,lastColumn).setValue('Pending');
  if(datestatus!=today){
    if(currentstatusExt.split("\"")[1]=="Service Issue" && Number(date.split(",")[0].split(" ")[1])!=Number(dateREPLACEBANK2.split(",")[0].split(": ")[1].split(" ")[1]))
  {
    extension(currentstatusExt.split("\"")[1]);
  }
    else if (Number(date.split(",")[0].split(" ")[1])==Number(dateREPLACEBANK2.split(",")[0].split(": ")[1].split(" ")[1]))
  {
     emailstatus(currentstatus.split("\"")[1],date,msg,today);
  }
  }  
  return;
}


function sendEmail(REPLACEBANK2,hourly)
{
  var REPLACEBANK2data=REPLACEBANK2+4;
  var date=SpreadsheetApp.getActiveSheet().getRange(1,1).getValue();
  var REPLACEBANK1 = SpreadsheetApp.getActiveSheet().getRange(5,2).getDisplayValue();
  var REPLACEBANK2 = SpreadsheetApp.getActiveSheet().getRange(46,2).getDisplayValue();
  Logger.log(REPLACEBANK2data);
  var Total = SpreadsheetApp.getActiveSheet().getRange(5,2).getValue()+SpreadsheetApp.getActiveSheet().getRange(46,2).getValue();
  //var Total = parseFloat(REPLACEBANK1.split("**")[REPLACEBANK1.split("**").length-1].split("$")[1].split(",").join(""))+parseFloat(REPLACEBANK2.split("**")[REPLACEBANK2.split("**").length-1].split("$")[1].split(",").join(""))
  var Curr = formatCurrency('$', Total);
  if (hourly =="17:01"||hourly =="17:02"||hourly =="17:03"||hourly =="17:04"||hourly =="17:05")
  {
  fivepm(date,REPLACEBANK2,REPLACEBANK1,Curr);
  }
  else if (hourly =="18:31"||hourly =="18:32"||hourly =="18:33")
  {
  sixthirty(date,REPLACEBANK2,REPLACEBANK1,Curr);
  }
  else if (hourly =="99:99")
  {
  sixthirty(date,REPLACEBANK2,REPLACEBANK1,Curr);
  }
  else
  {
  composeApprovedEmail(date,REPLACEBANK2,REPLACEBANK1,Curr);
  }
}

function formatCurrency(symbol, amount) {
  var aDigits = amount.toFixed(2).split(".");
  aDigits[0] = aDigits[0].split("").reverse().join("")
    .replace(/(\d{3})(?=\d)/g,"$1,").split("").reverse().join("");
  return symbol + aDigits.join(".");
}
 
 
function composeApprovedEmail(dt,REPLACEBANK2,REPLACEBANK1,total)
{

  var email = "REPLACE";
  //var email = "REPLACE";
 
  var message = " \n"+dt+" \n"+" \n"+"REPLACEBANK1:  "+REPLACEBANK1+" \n"+" \n"+"REPLACEBANK2: "+REPLACEBANK2+" \n"+" \n"+"Total Fedline Balance: "+total;

  var subject = "Fedline Balance Update";
 
  MailApp.sendEmail(email, subject, message);

}
function fivepm(dt,REPLACEBANK2,REPLACEBANK1,total)
{

  var email = "REPLACE";
 
  var message = " \n"+dt+" \n"+" \n"+"REPLACEBANK1:  "+REPLACEBANK1+" \n"+" \n"+"REPLACEBANK2: "+REPLACEBANK2+" \n"+" \n"+"Total Fedline Balance: "+total;

  var subject = "5pm Fedline Balance";
 
  MailApp.sendEmail(email, subject, message);
}
function sixthirty(dt,REPLACEBANK2,REPLACEBANK1,total)
{
  const sheet = SpreadsheetApp.openById("1fsTGPqMDzsELl4qlcpXIkHkNOL0rmVewt05LNIiZ5Ls").getSheets()[0];
  var REPLACEBANK1numb = parseFloat(REPLACEBANK1.split("**")[REPLACEBANK1.split("**").length-1].split("$")[1].split(",").join(""))
  var REPLACEBANK2numb = parseFloat(REPLACEBANK2.split("**")[REPLACEBANK2.split("**").length-1].split("$")[1].split(",").join(""))
  sheet.getRange(12,3).setValue(REPLACEBANK1numb);
  sheet.getRange(13,3).setValue(REPLACEBANK2numb);
  var email = "REPLACE";
 
  var message = " \n"+dt+" \n"+" \n"+"REPLACEBANK1:  "+REPLACEBANK1+" \n"+" \n"+"REPLACEBANK2: "+REPLACEBANK2+" \n"+" \n"+"Total Fedline Balance: "+total;

  var subject = "Closing Fedline Balance";
 
  MailApp.sendEmail(email, subject, message);
}
function emailme(today)
{
  var REPLACEBANK1 = SpreadsheetApp.getActiveSheet().getRange(17,6).getDisplayValue();
  var REPLACEBANK2 = SpreadsheetApp.getActiveSheet().getRange(58,6).getDisplayValue();
  var email = "REPLACE";
 
  var message = "ACH, check and cash letter activity"+" \n"+"REPLACEBANK1:  "+REPLACEBANK1+" \n"+"REPLACEBANK2: "+REPLACEBANK2;

  var subject = "ACH & Cash Balances for "+today;
 
  MailApp.sendEmail(email, subject, message);

}
function emailstatus(status,date,msg,today)
{
 
  var ss = SpreadsheetApp.openById("1xz1hheHURpwkRojtyfbz02YVbpXmIh9pz_Ns8YfICdg");
  var sheets = ss.getSheets();
  var datestatus = sheets[0].getRange(2,2).setValue(today);
  var email = "REPLACE";
 
  var message = ""+date+" \n"+"Status:  "+status+" \n"+msg;

  var subject = "Fedline is "+status+" for "+date;
 
  MailApp.sendEmail(email, subject, message);
  hourly="99:99"
  EOD(hourly)
}
function extension(status,date,msg)
{
 
  var email = "REPLACE";
 
  var message = "Please check the current Status here: https://www.frbservices.org/app/status/serviceStatus.do";

  var subject = "Fedwire Funds Service has issued an EXTENSION";
 
  MailApp.sendEmail(email, subject, message);

}


function convert(today){
 
 var sourceFolderId = "1YUt_hTNClRUzhzrOOCPKas6_mkopTMRY"; // Folder ID including source files.
  var destinationFolderId = "1YUt_hTNClRUzhzrOOCPKas6_mkopTMRY"; // Folder ID that the converted files are put.

  var getFileIds = function (folder, fileList, q) {
    var files = folder.searchFiles(q);
    while (files.hasNext()) {
      var f = files.next();
      fileList.push({id: f.getId(), fileName: f.getName().split(".")[0].trim()});
    }
    var folders = folder.getFolders();
    while (folders.hasNext()) getFileIds(folders.next(), fileList, q);
    return fileList;
  };
  var sourceFiles = getFileIds(DriveApp.getFolderById(sourceFolderId), [], "mimeType='" + MimeType.MICROSOFT_EXCEL + "' or mimeType='" + MimeType.MICROSOFT_EXCEL_LEGACY + "'");
  var destinationFiles = getFileIds(DriveApp.getFolderById(destinationFolderId), [], "mimeType='" + MimeType.GOOGLE_SHEETS + "'");
  var createFiles = sourceFiles.filter(function(e) {return destinationFiles.every(function(f) {return f.fileName !== e.fileName});});
  var updateFiles = sourceFiles.reduce(function(ar, e) {
    var dst = destinationFiles.filter(function(f) {return f.fileName === e.fileName});
    if (dst.length > 0) {
      e.to = dst[0].id;
      ar.push(e);
    }
    return ar;
  }, []);
  if (createFiles.length > 0) createFiles.forEach(function(e) {Drive.Files.insert({mimeType: MimeType.GOOGLE_SHEETS, parents: [{id: destinationFolderId}], title: e.fileName}, DriveApp.getFileById(e.id))});
  if (updateFiles.length > 0) updateFiles.forEach(function(e) {Drive.Files.update({}, e.to, DriveApp.getFileById(e.id))});
  sendMail(updateFiles[0].to,today);
}
function expectedpayment(today){
 var sourceFolderId = "1s1neN4W8DHQY-8tBqlDj7n69-YHtgNCT"; // Folder ID including source files.
  var destinationFolderId = "1s1neN4W8DHQY-8tBqlDj7n69-YHtgNCT"; // Folder ID that the converted files are put.

  var getFileIds = function (folder, fileList, q) {
    var files = folder.searchFiles(q);
    while (files.hasNext()) {
      var f = files.next();
      fileList.push({id: f.getId(), fileName: f.getName().split(".")[0].trim()});
    }
    var folders = folder.getFolders();
    while (folders.hasNext()) getFileIds(folders.next(), fileList, q);
    return fileList;
  };
  var sourceFiles = getFileIds(DriveApp.getFolderById(sourceFolderId), [], "mimeType='" + MimeType.MICROSOFT_EXCEL + "' or mimeType='" + MimeType.MICROSOFT_EXCEL_LEGACY + "'");
  var destinationFiles = getFileIds(DriveApp.getFolderById(destinationFolderId), [], "mimeType='" + MimeType.GOOGLE_SHEETS + "'");
  var createFiles = sourceFiles.filter(function(e) {return destinationFiles.every(function(f) {return f.fileName !== e.fileName});});
  var updateFiles = sourceFiles.reduce(function(ar, e) {
    var dst = destinationFiles.filter(function(f) {return f.fileName === e.fileName});
    if (dst.length > 0) {
      e.to = dst[0].id;
      ar.push(e);
    }
    return ar;
  }, []);
  if (createFiles.length > 0) createFiles.forEach(function(e) {Drive.Files.insert({mimeType: MimeType.GOOGLE_SHEETS, parents: [{id: destinationFolderId}], title: e.fileName}, DriveApp.getFileById(e.id))});
  if (updateFiles.length > 0) updateFiles.forEach(function(e) {Drive.Files.update({}, e.to, DriveApp.getFileById(e.id))});
  expectedpaymenttable(updateFiles[0].to,today);
 
}



function expectedpaymenttable(id,today){
 var sheet = SpreadsheetApp.openById(id);
 var ss   = sheet.getSheetByName("Sheet1");
 var startingRow;
 var endingRow;
 for(var b = 4; b <= ss.getLastRow(); b++) {
          var check = ss.getRange(b,1).getDisplayValue();
   
        var ending = today+' Total';
       
        if (check === today)
        {
          startingRow=b;
        }
        else if (check === ending)
        {
          endingRow=b;
        }
      }
  var range = ss.getRange(startingRow,1,endingRow-startingRow+1,ss.getLastColumn());
  var range2 = ss.getRange(4,1,1,ss.getLastColumn());
  var data = range2.getDisplayValues().concat(range.getDisplayValues());

var TABLEFORMAT = 'cellspacing="2" cellpadding="2" dir="ltr" border="1" style="width:100%;table-layout:fixed;font-size:10pt;font-family:arial,sans,sans-serif;border-collapse:collapse;border:1px solid #ccc;font-weight:normal;color:black;background-color:white;text-align:center;text-decoration:none;font-style:normal;'
var htmltable = '<table ' + TABLEFORMAT +' ">';

for (row = 0; row<data.length; row++){

htmltable += '<tr>';

for (col = 0 ;col<data[row].length; col++){
  if (data[row][col] === "" || 0) {htmltable += '<td>' + ' ' + '</td>';}
  else
    if (row === 0)  {
      htmltable += '<th>' + data[row][col] + '</th>';
    }

  else {htmltable += '<td>' + data[row][col] + '</td>';}
}

     htmltable += '</tr>';
}

     htmltable += '</table>';
     //Logger.log(data);
     //Logger.log(htmltable);
     var email = "REPLACE";
MailApp.sendEmail(email, 'Expected Payment '+today,'' ,{htmlBody: htmltable})
}


function dashboard(){
 var sourceFolderId = "1i7O50pLHwpIk5YFh8o-ECxhQQzajwUmd"; // Folder ID including source files.
  var destinationFolderId = "1wTHqTwHH5bnrJTqjPPfPEiq2bca_rHv8"; // Folder ID that the converted files are put.

  var getFileIds = function (folder, fileList, q) {
    var files = folder.searchFiles(q);
    while (files.hasNext()) {
      var f = files.next();
      fileList.push({id: f.getId(), fileName: f.getName().split(".")[0].trim()});
    }
    var folders = folder.getFolders();
    while (folders.hasNext()) getFileIds(folders.next(), fileList, q);
    return fileList;
  };
  var sourceFiles = getFileIds(DriveApp.getFolderById(sourceFolderId), [], "mimeType='" + MimeType.MICROSOFT_EXCEL + "' or mimeType='" + MimeType.MICROSOFT_EXCEL_LEGACY + "'");
  var destinationFiles = getFileIds(DriveApp.getFolderById(destinationFolderId), [], "mimeType='" + MimeType.GOOGLE_SHEETS + "'");
  var createFiles = sourceFiles.filter(function(e) {return destinationFiles.every(function(f) {return f.fileName !== e.fileName});});
  var updateFiles = sourceFiles.reduce(function(ar, e) {
    var dst = destinationFiles.filter(function(f) {return f.fileName === e.fileName});
    if (dst.length > 0) {
      e.to = dst[0].id;
      ar.push(e);
    }
    return ar;
  }, []);
  if (createFiles.length > 0) createFiles.forEach(function(e) {Drive.Files.insert({mimeType: MimeType.GOOGLE_SHEETS, parents: [{id: destinationFolderId}], title: e.fileName}, DriveApp.getFileById(e.id))});
  if (updateFiles.length > 0) updateFiles.forEach(function(e) {Drive.Files.update({}, e.to, DriveApp.getFileById(e.id))});
 
  var ss = SpreadsheetApp.openById("1fsTGPqMDzsELl4qlcpXIkHkNOL0rmVewt05LNIiZ5Ls");
  var statussheet = ss.getSheets();
  var BMOBalance;
  var REPLACEBANKALLBalance;
  var BMOQuantum;
  var BMOTime;
  var InterTime;
  var REPLACEBANK2Time;
  var REPLACEBANK2Quantum;
  var BalanceDate;
  var REPLACEBANK1Ending;
  var REPLACEBANK2Ending;
  var EuroBalance;
  var Interdate;
  var GBPBalance;
  var CollateralBalance;
  for(var i = 0; i <=5; i++) {
  var currentfile = updateFiles[i].to;
  var opensheet = SpreadsheetApp.openById(currentfile);
  var currentsheet = opensheet.getSheets();
  var currenttitle = currentsheet[0].getRange(1,1).getValues();
  var currenttitle2 = currentsheet[0].getRange(1,4).getValues();
 
  var BMOrow;
  var REPLACEBANK2row;
  var REPLACEBANK1EndingRow;
  var REPLACEBANK2EndingRow;
  var EuroRow;
  var GBPRow;
  var CollateralRow;
    if(currenttitle == ""&& currenttitle2 == ""){
      for(var b = 15; b <= currentsheet[0].getLastRow(); b++) {
          var check = currentsheet[0].getRange(b,2).getDisplayValue();
        //Logger.log(check);
        //var ending = today+' Total';
       
        if (check == "00021597213")
        {
          BMOrow=b;
        }
        else if (check == "5720081131")
        {
          REPLACEBANK2row=b;
        }
        else if (check == "051405515")
        {
          REPLACEBANK1EndingRow=b;
        }
        else if (check == "056073502")
        {
          REPLACEBANK2EndingRow=b;
        }
        else if (check == "8033045264")
        {
          EuroRow=b;
        }
        else if (check == "2081347274")
        {
          CollateralRow=b;
        }
        else if (check == "20325383386619")
        {
          GBPRow=b;
        }
       
      }
    currenttitle=currentsheet[0].getRange(2,21).getValues();
    BMOBalance=currentsheet[0].getRange(BMOrow,28).getValues();
    REPLACEBANKALLBalance=currentsheet[0].getRange(REPLACEBANK2row,28).getValues();
    REPLACEBANK1Ending=currentsheet[0].getRange(REPLACEBANK1EndingRow,28).getValues();
    REPLACEBANK2Ending= currentsheet[0].getRange(REPLACEBANK2EndingRow,28).getValues();
    EuroBalance= currentsheet[0].getRange(EuroRow,28).getValues();
    CollateralBalance= currentsheet[0].getRange(CollateralRow,28).getValues();
    try {
    GBPBalance= currentsheet[0].getRange(GBPRow,28).getValues();
    } catch (e) {
    Logger.log("GBP Bank Holiday")
    }
    BalanceDate=currentsheet[0].getRange(4,4).getValues();
    }
    else if (currenttitle == "PD Recon - Updated – Prior Day Match - AccountsSummary-All Accounts")
    {
      statussheet[0].getRange(2,8).setValue("TRUE");
      for(var b = 5; b <= currentsheet[0].getLastRow(); b++) {
          var check = currentsheet[0].getRange(b,1).getValues();
        if (check != "Detail Matched" && check != "Fully Reconciled / Marked")
        {
          statussheet[0].getRange(2,8).setValue("FALSE");
          break;
        }
      }
    }
    else if(currenttitle == "CashXplorer-Bank of Montreal - CD BMO")
    {
      BMOQuantum = currentsheet[0].getRange(currentsheet[0].getLastRow()-2,currentsheet[0].getLastColumn()).getValues();
      BMOTime= currentsheet[0].getRange(2,1).getValues();
      statussheet[0].getRange(3,2).setValue(BMOTime);
      statussheet[0].getRange(3,3).setValue(BMOQuantum);
    }
    else if(currenttitle == "CashXplorer-Intercompany Loans - Intercompany- US")
    {
      InterTime= currentsheet[0].getRange(2,1).getValues();
      statussheet[0].getRange(7,2).setValue(InterTime);
      for(var b = 5; b <= currentsheet[0].getLastRow(); b++) {
          var check = currentsheet[0].getRange(b,2).getDisplayValue();      
        if (check == "UL 10000-11000, UL 10000(REPLACEBANKALL) loans to 11000(REPLACEBANK1)")
        {
          statussheet[0].getRange(8,2).setValue(currentsheet[0].getRange(b,4).getDisplayValue());
          statussheet[0].getRange(8,3).setValue(currentsheet[0].getRange(b,5).getDisplayValue());
          statussheet[0].getRange(8,4).setValue(currentsheet[0].getRange(b,6).getDisplayValue());
          statussheet[0].getRange(8,5).setValue(currentsheet[0].getRange(b,7).getDisplayValue());
        }
        else if (check == "UL 10000-21100, UL 10000(REPLACEBANKALL) loans to 21100(REPLACEBANK2)")
        {
          statussheet[0].getRange(9,2).setValue(currentsheet[0].getRange(b,4).getDisplayValue());
          statussheet[0].getRange(9,3).setValue(currentsheet[0].getRange(b,5).getDisplayValue());
          statussheet[0].getRange(9,4).setValue(currentsheet[0].getRange(b,6).getDisplayValue());
          statussheet[0].getRange(9,5).setValue(currentsheet[0].getRange(b,7).getDisplayValue());
        }
        else if (check == "UL 21100-11000, UL 21100(REPLACEBANK2) loans to 11000(REPLACEBANK1)"){
          statussheet[0].getRange(10,2).setValue(currentsheet[0].getRange(b,4).getDisplayValue());
          statussheet[0].getRange(10,3).setValue(currentsheet[0].getRange(b,5).getDisplayValue());
          statussheet[0].getRange(10,4).setValue(currentsheet[0].getRange(b,6).getDisplayValue());
          statussheet[0].getRange(10,5).setValue(currentsheet[0].getRange(b,7).getDisplayValue());
        }}
    }
    else if(currenttitle == "CashXplorer-REPLACEBANK2 DDAs - CD Concentration Funding")
    {
      REPLACEBANK2Quantum = currentsheet[0].getRange(currentsheet[0].getLastRow()-2,currentsheet[0].getLastColumn()-6).getValues();
      REPLACEBANK2Time= currentsheet[0].getRange(2,1).getValues();
      statussheet[0].getRange(4,2).setValue(REPLACEBANK2Time);
      statussheet[0].getRange(4,3).setValue(REPLACEBANK2Quantum);
    }
    else if(currenttitle2="SecType")
    {
      Interdate=currentsheet[0].getRange(2,2).getValues();
    }
  Logger.log(currenttitle);
   }
  Logger.log(Interdate);
  statussheet[0].getRange(3,4).setValue(BMOBalance);
  statussheet[0].getRange(4,4).setValue(REPLACEBANKALLBalance);
  statussheet[0].getRange(5,2).setValue(BalanceDate);
  statussheet[0].getRange(6,2).setValue(Interdate);
  statussheet[0].getRange(12,4).setValue(REPLACEBANK1Ending);
  statussheet[0].getRange(13,4).setValue(REPLACEBANK2Ending);
  statussheet[0].getRange(14,4).setValue(EuroBalance);
  statussheet[0].getRange(16,4).setValue(GBPBalance);
  statussheet[0].getRange(15,4).setValue(CollateralBalance);
  const sheet = SpreadsheetApp.openById('1tF-wGPJ-Sg_Dw_9wf49YYs4AvXP8szpj6tCndLMOKUE').getSheets()[0];
  var yesterday=sheet.getRange(6,3).getDisplayValues();
  var suspensedate=sheet.getRange(6,1).getValues();
  statussheet[0].getRange(11,2).setValue(suspensedate);
  statussheet[0].getRange(11,8).setValue(yesterday);
}

function sendMail(id,today){
 var sh = SpreadsheetApp.openById(id);
 var data = sh.getDataRange().getDisplayValues();
  //var htmltable =[];

var TABLEFORMAT = 'cellspacing="2" cellpadding="2" dir="ltr" border="1" style="width:100%;table-layout:fixed;font-size:10pt;font-family:arial,sans,sans-serif;border-collapse:collapse;border:1px solid #ccc;font-weight:normal;color:black;background-color:white;text-align:center;text-decoration:none;font-style:normal;'
var htmltable = '<table ' + TABLEFORMAT +' ">';

for (row = 0; row<data.length; row++){

htmltable += '<tr>';

for (col = 0 ;col<data[row].length; col++){
  if (data[row][col] === "" || 0) {htmltable += '<td>' + ' ' + '</td>';}
  else
    if (row === 0)  {
      htmltable += '<th>' + data[row][col] + '</th>';
    }

  else {htmltable += '<td>' + data[row][col] + '</td>';}
}

     htmltable += '</tr>';
}

     htmltable += '</table>';
     Logger.log(data);
     Logger.log(htmltable);
     var email = "REPLACE";
MailApp.sendEmail(email, 'CAD Balances for '+today,'' ,{htmlBody: htmltable})
}

function BankRecon(){
var threads = GmailApp.search("QTM PROD-Quantum Bank Account Balance Report");// from today
var messages = threads[0].getMessages();
var message = messages[messages.length - 1];
var attachment = message.getAttachments()[0];
attachment.setContentType('application/vnd.ms-excel');
//var attachmentBlob = attachment.copyBlob();
//  var file = DriveApp.createFile(attachmentBlob);
//  Drive.Files.insert(
//          {
//            title: 'Bank Account Report',
//            mimeType: attachment.getContentType(),
//            parents: [{ id: '1i7O50pLHwpIk5YFh8o-ECxhQQzajwUmd' }]
//          },
//     attachment.copyBlob())
  Drive.Files.update({
    title: 'Bank Account Report', mimeType: attachment.getContentType()
  }, '1SadAL7Ulx5Ry7B1b8uiO_F2XDBNeNN0r', attachment.copyBlob());
 
  var threads = GmailApp.search("[External Sender] Interco Loan Report");// from today
var messages = threads[0].getMessages();
var message = messages[messages.length - 1];
var attachment = message.getAttachments()[0];
attachment.setContentType('application/vnd.ms-excel');
//var attachmentBlob = attachment.copyBlob();
//  var file = DriveApp.createFile(attachmentBlob);
//  Drive.Files.insert(
//          {
//            title: 'Bank Account Report',
//            mimeType: attachment.getContentType(),
//            parents: [{ id: '1i7O50pLHwpIk5YFh8o-ECxhQQzajwUmd' }]
//          },
//     attachment.copyBlob())
  Drive.Files.update({
    title: 'Interco Loan Report', mimeType: attachment.getContentType()
  }, '146NF9c2_QuK61nc2e6Ad0G_jmWb9V1Jo', attachment.copyBlob());
//}
 
            }

function getTextFromPDF() {
  var threads = GmailApp.search("QTM PROD-Quantum Suspense Report ");// from today
var messages = threads[0].getMessages();
var message = messages[messages.length - 1];
var attachment = message.getAttachments()[0];
attachment.setContentType('application/pdf');
  Drive.Files.update({
    title: 'Suspense Report', mimeType: attachment.getContentType()
  }, '1BdpstldzEVJHirbElIjiEbrto7iKrn-U', attachment.copyBlob());
  //var fileID='1BdpstldzEVJHirbElIjiEbrto7iKrn-U';
  var blob = DriveApp.getFileById('1BdpstldzEVJHirbElIjiEbrto7iKrn-U').getBlob()
  var resource = {
    title: blob.getName(),
    mimeType: blob.getContentType()
  };
  var options = {
  ocr: true,
    ocrLanguage: "en"
  };
   //Convert the pdf to a Google Doc with ocr.
  var file = Drive.Files.insert(resource, blob, options);
 
   //Get the texts from the newly created text.
  var doc = DocumentApp.openById(file.id);
  var text = doc.getBody().getText();
  Drive.Files.remove(doc.getId());
  var lines = text.split("\n")
 
  const sheet = SpreadsheetApp.openById('1tF-wGPJ-Sg_Dw_9wf49YYs4AvXP8szpj6tCndLMOKUE').getSheets()[0];
  //Logger.log(lines.length)
  for (i=0;i<lines.length;i++){
    sheet.getRange(i+1,1).setValue(lines[i]);
  }
  var today=sheet.getRange(6,2).getValues();
  if(lines.length>10 && today){
    suspensereport()
    return;
  }
}
  function suspensereport()
{
 
  var email = "function onEdit() {
  var file = DriveApp.getFilesByName("REPLACEBANK2Balance.csv").next();
  var date= file.getLastUpdated();
  var current= new Date()
  var difference=(current.getTime()-date.getTime())/60000
  var fedline = DriveApp.getFilesByName("FedlineStatus.csv").next();
  var fed= fedline.getLastUpdated();
  var cad = DriveApp.getFileById('1Dx-IW6SCghxgRetyePEsIzKlTaaZ-U_e').getLastUpdated();
  var hourly = Utilities.formatDate(new Date(),"GMT-4:00", "HH:mm");
  var lastdate = Utilities.formatDate(date,"GMT-4:00", "MM/dd/yyyy");
  var caddate = Utilities.formatDate(cad,"GMT-4:00", "MM/dd/yyyy");
  var feddate = Utilities.formatDate(fed,"GMT-4:00", "MM/dd/yyyy");
  var today = Utilities.formatDate(new Date(),"GMT-4:00", "MM/dd/yyyy");
  var todayformat = Utilities.formatDate(new Date(),"GMT-4:00", "M/d/yyyy");
  if(difference>1.006)
  {
    if (hourly =="09:08" & lastdate==today)
  {
    emailme(today);
  }
    else if (hourly =="14:05" & caddate==today)
  {
    convert(today);
  }
    else if (hourly =="17:18")
  {
    getTextFromPDF();
  }
    else if ((hourly =="18:03"||hourly =="18:18"||hourly =="18:33"||hourly =="18:48"||hourly =="19:03"||hourly =="19:18"||hourly =="19:33"||hourly =="19:48"||hourly =="20:03"||hourly =="20:18"||hourly =="20:33"||hourly =="20:48")& feddate==today)
  {
    FedlineStatus(today);
  }
  else if (hourly =="08:48")
  {
    BankRecon();
  }
  else if (hourly =="08:58")
  {
    dashboard();
  }
  else if (hourly =="07:58")
  {
    expectedpayment(todayformat);
  }
    Logger.log(difference);
    return;
  }
  EOD(hourly)
  }

function QuantumMatch(){
 
 var sourceFolderId = "1w6uHIqvHkGI0sid7vkAjpc0S2hBlc1SY"; // Folder ID including source files.
  var destinationFolderId = "1w6uHIqvHkGI0sid7vkAjpc0S2hBlc1SY"; // Folder ID that the converted files are put.

  var getFileIds = function (folder, fileList, q) {
    var files = folder.searchFiles(q);
    while (files.hasNext()) {
      var f = files.next();
      fileList.push({id: f.getId(), fileName: f.getName().split(".")[0].trim()});
    }
    var folders = folder.getFolders();
    while (folders.hasNext()) getFileIds(folders.next(), fileList, q);
    return fileList;
  };
  var sourceFiles = getFileIds(DriveApp.getFolderById(sourceFolderId), [], "mimeType='" + MimeType.MICROSOFT_EXCEL + "' or mimeType='" + MimeType.MICROSOFT_EXCEL_LEGACY + "'");
  var destinationFiles = getFileIds(DriveApp.getFolderById(destinationFolderId), [], "mimeType='" + MimeType.GOOGLE_SHEETS + "'");
  var createFiles = sourceFiles.filter(function(e) {return destinationFiles.every(function(f) {return f.fileName !== e.fileName});});
  var updateFiles = sourceFiles.reduce(function(ar, e) {
    var dst = destinationFiles.filter(function(f) {return f.fileName === e.fileName});
    if (dst.length > 0) {
      e.to = dst[0].id;
      ar.push(e);
    }
    return ar;
  }, []);
  if (createFiles.length > 0) createFiles.forEach(function(e) {Drive.Files.insert({mimeType: MimeType.GOOGLE_SHEETS, parents: [{id: destinationFolderId}], title: e.fileName}, DriveApp.getFileById(e.id))});
  if (updateFiles.length > 0) updateFiles.forEach(function(e) {Drive.Files.update({}, e.to, DriveApp.getFileById(e.id))});
 var sh = SpreadsheetApp.openById(updateFiles[0].to).getSheets()[0];
 var data = sh.getDataRange().getDisplayValues();
 var sheet = SpreadsheetApp.openById('1haf4IYeIfUJL8ClNzPjSxMtTQp32IAKTSl1SXQO0DHo').getSheets()[0];
 sheet.clear()
 sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
 var sh1 = SpreadsheetApp.openById(updateFiles[0].to).getSheets()[1];
 var data1 = sh1.getDataRange().getDisplayValues();
 var sheet1 = SpreadsheetApp.openById('1w7qUXwp4cBxxGncWYrmslEpYpSqjV0z9P7XKlw2u8F4').getSheets()[0];
 sheet1.clear()
 sheet1.getRange(1, 1, data1.length, data1[0].length).setValues(data1);
}

function EOD(hourly) {
  //CSV to Sheet
  var file = DriveApp.getFilesByName("REPLACEBANK1Balance.csv").next();
  var date= file.getLastUpdated()
  var REPLACEBANK1 = Utilities.parseCsv(file.getBlob().getDataAsString());
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.clear()
  sheet.getRange(1, 1, REPLACEBANK1.length, REPLACEBANK1[0].length).setValues(REPLACEBANK1);
  //sheet.insertRowsAfter(REPLACEBANK1.length+1, 1);
  var file1 = DriveApp.getFilesByName("REPLACEBANK2Balance.csv").next();
  var REPLACEBANK2 = Utilities.parseCsv(file1.getBlob().getDataAsString());
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(sheet.getLastRow()+1,1,REPLACEBANK2.length,REPLACEBANK2[0].length).setValues(REPLACEBANK2);
  sendEmail(REPLACEBANK2.length,hourly);
}

function FedlineStatus(today) {
  //CSV to Sheet
  var file = DriveApp.getFilesByName("FedlineStatus.csv").next();
  var Status = Utilities.parseCsv(file.getBlob().getDataAsString());
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(100, 1, Status.length, Status[0].length).setValues(Status);
  var currentstatus= SpreadsheetApp.getActiveSheet().getRange(105,2).getDisplayValue();
  var currentstatusExt= SpreadsheetApp.getActiveSheet().getRange(104,2).getDisplayValue();
  var date= SpreadsheetApp.getActiveSheet().getRange(108,3).getDisplayValue();
  var dateREPLACEBANK2= SpreadsheetApp.getActiveSheet().getRange(1,1).getDisplayValue();
  var msg= SpreadsheetApp.getActiveSheet().getRange(110,1).getDisplayValue();
  //Logger.log(currentstatusExt.split("\"")[1])
  var ss = SpreadsheetApp.openById("1xz1hheHURpwkRojtyfbz02YVbpXmIh9pz_Ns8YfICdg");
  var sheets = ss.getSheets();
  var datestatus = sheets[0].getRange(2,2).getDisplayValue();
  Logger.log(datestatus)
  Logger.log(today)
  //sheet.getRange(lastRow,lastColumn).setValue('Pending');
  if(datestatus!=today){
    if(currentstatusExt.split("\"")[1]=="Service Issue" && Number(date.split(",")[0].split(" ")[1])!=Number(dateREPLACEBANK2.split(",")[0].split(": ")[1].split(" ")[1]))
  {
    extension(currentstatusExt.split("\"")[1]);
  }
    else if (Number(date.split(",")[0].split(" ")[1])==Number(dateREPLACEBANK2.split(",")[0].split(": ")[1].split(" ")[1]))
  {
     emailstatus(currentstatus.split("\"")[1],date,msg,today);
  }
  }  
  return;
}


function sendEmail(REPLACEBANK2,hourly)
{
  var REPLACEBANK2data=REPLACEBANK2+4;
  var date=SpreadsheetApp.getActiveSheet().getRange(1,1).getValue();
  var REPLACEBANK1 = SpreadsheetApp.getActiveSheet().getRange(5,2).getDisplayValue();
  var REPLACEBANK2 = SpreadsheetApp.getActiveSheet().getRange(46,2).getDisplayValue();
  Logger.log(REPLACEBANK2data);
  var Total = SpreadsheetApp.getActiveSheet().getRange(5,2).getValue()+SpreadsheetApp.getActiveSheet().getRange(46,2).getValue();
  //var Total = parseFloat(REPLACEBANK1.split("**")[REPLACEBANK1.split("**").length-1].split("$")[1].split(",").join(""))+parseFloat(REPLACEBANK2.split("**")[REPLACEBANK2.split("**").length-1].split("$")[1].split(",").join(""))
  var Curr = formatCurrency('$', Total);
  if (hourly =="17:01"||hourly =="17:02"||hourly =="17:03"||hourly =="17:04"||hourly =="17:05")
  {
  fivepm(date,REPLACEBANK2,REPLACEBANK1,Curr);
  }
  else if (hourly =="18:31"||hourly =="18:32"||hourly =="18:33")
  {
  sixthirty(date,REPLACEBANK2,REPLACEBANK1,Curr);
  }
  else if (hourly =="99:99")
  {
  sixthirty(date,REPLACEBANK2,REPLACEBANK1,Curr);
  }
  else
  {
  composeApprovedEmail(date,REPLACEBANK2,REPLACEBANK1,Curr);
  }
}

function formatCurrency(symbol, amount) {
  var aDigits = amount.toFixed(2).split(".");
  aDigits[0] = aDigits[0].split("").reverse().join("")
    .replace(/(\d{3})(?=\d)/g,"$1,").split("").reverse().join("");
  return symbol + aDigits.join(".");
}
 
 
function composeApprovedEmail(dt,REPLACEBANK2,REPLACEBANK1,total)
{

  var email = "REPLACE";
  //var email = "REPLACE";
 
  var message = " \n"+dt+" \n"+" \n"+"REPLACEBANK1:  "+REPLACEBANK1+" \n"+" \n"+"REPLACEBANK2: "+REPLACEBANK2+" \n"+" \n"+"Total Fedline Balance: "+total;

  var subject = "Fedline Balance Update";
 
  MailApp.sendEmail(email, subject, message);

}
function fivepm(dt,REPLACEBANK2,REPLACEBANK1,total)
{

  var email = "REPLACE";
 
  var message = " \n"+dt+" \n"+" \n"+"REPLACEBANK1:  "+REPLACEBANK1+" \n"+" \n"+"REPLACEBANK2: "+REPLACEBANK2+" \n"+" \n"+"Total Fedline Balance: "+total;

  var subject = "5pm Fedline Balance";
 
  MailApp.sendEmail(email, subject, message);
}
function sixthirty(dt,REPLACEBANK2,REPLACEBANK1,total)
{
  const sheet = SpreadsheetApp.openById("1fsTGPqMDzsELl4qlcpXIkHkNOL0rmVewt05LNIiZ5Ls").getSheets()[0];
  var REPLACEBANK1numb = parseFloat(REPLACEBANK1.split("**")[REPLACEBANK1.split("**").length-1].split("$")[1].split(",").join(""))
  var REPLACEBANK2numb = parseFloat(REPLACEBANK2.split("**")[REPLACEBANK2.split("**").length-1].split("$")[1].split(",").join(""))
  sheet.getRange(12,3).setValue(REPLACEBANK1numb);
  sheet.getRange(13,3).setValue(REPLACEBANK2numb);
  var email = "REPLACE";
 
  var message = " \n"+dt+" \n"+" \n"+"REPLACEBANK1:  "+REPLACEBANK1+" \n"+" \n"+"REPLACEBANK2: "+REPLACEBANK2+" \n"+" \n"+"Total Fedline Balance: "+total;

  var subject = "Closing Fedline Balance";
 
  MailApp.sendEmail(email, subject, message);
}
function emailme(today)
{
  var REPLACEBANK1 = SpreadsheetApp.getActiveSheet().getRange(17,6).getDisplayValue();
  var REPLACEBANK2 = SpreadsheetApp.getActiveSheet().getRange(58,6).getDisplayValue();
  var email = "REPLACE";
 
  var message = "ACH, check and cash letter activity"+" \n"+"REPLACEBANK1:  "+REPLACEBANK1+" \n"+"REPLACEBANK2: "+REPLACEBANK2;

  var subject = "ACH & Cash Balances for "+today;
 
  MailApp.sendEmail(email, subject, message);

}
function emailstatus(status,date,msg,today)
{
 
  var ss = SpreadsheetApp.openById("1xz1hheHURpwkRojtyfbz02YVbpXmIh9pz_Ns8YfICdg");
  var sheets = ss.getSheets();
  var datestatus = sheets[0].getRange(2,2).setValue(today);
  var email = "REPLACE";
 
  var message = ""+date+" \n"+"Status:  "+status+" \n"+msg;

  var subject = "Fedline is "+status+" for "+date;
 
  MailApp.sendEmail(email, subject, message);
  hourly="99:99"
  EOD(hourly)
}
function extension(status,date,msg)
{
 
  var email = "REPLACE";
 
  var message = "Please check the current Status here: https://www.frbservices.org/app/status/serviceStatus.do";

  var subject = "Fedwire Funds Service has issued an EXTENSION";
 
  MailApp.sendEmail(email, subject, message);

}


function convert(today){
 
 var sourceFolderId = "1YUt_hTNClRUzhzrOOCPKas6_mkopTMRY"; // Folder ID including source files.
  var destinationFolderId = "1YUt_hTNClRUzhzrOOCPKas6_mkopTMRY"; // Folder ID that the converted files are put.

  var getFileIds = function (folder, fileList, q) {
    var files = folder.searchFiles(q);
    while (files.hasNext()) {
      var f = files.next();
      fileList.push({id: f.getId(), fileName: f.getName().split(".")[0].trim()});
    }
    var folders = folder.getFolders();
    while (folders.hasNext()) getFileIds(folders.next(), fileList, q);
    return fileList;
  };
  var sourceFiles = getFileIds(DriveApp.getFolderById(sourceFolderId), [], "mimeType='" + MimeType.MICROSOFT_EXCEL + "' or mimeType='" + MimeType.MICROSOFT_EXCEL_LEGACY + "'");
  var destinationFiles = getFileIds(DriveApp.getFolderById(destinationFolderId), [], "mimeType='" + MimeType.GOOGLE_SHEETS + "'");
  var createFiles = sourceFiles.filter(function(e) {return destinationFiles.every(function(f) {return f.fileName !== e.fileName});});
  var updateFiles = sourceFiles.reduce(function(ar, e) {
    var dst = destinationFiles.filter(function(f) {return f.fileName === e.fileName});
    if (dst.length > 0) {
      e.to = dst[0].id;
      ar.push(e);
    }
    return ar;
  }, []);
  if (createFiles.length > 0) createFiles.forEach(function(e) {Drive.Files.insert({mimeType: MimeType.GOOGLE_SHEETS, parents: [{id: destinationFolderId}], title: e.fileName}, DriveApp.getFileById(e.id))});
  if (updateFiles.length > 0) updateFiles.forEach(function(e) {Drive.Files.update({}, e.to, DriveApp.getFileById(e.id))});
  sendMail(updateFiles[0].to,today);
}
function expectedpayment(today){
 var sourceFolderId = "1s1neN4W8DHQY-8tBqlDj7n69-YHtgNCT"; // Folder ID including source files.
  var destinationFolderId = "1s1neN4W8DHQY-8tBqlDj7n69-YHtgNCT"; // Folder ID that the converted files are put.

  var getFileIds = function (folder, fileList, q) {
    var files = folder.searchFiles(q);
    while (files.hasNext()) {
      var f = files.next();
      fileList.push({id: f.getId(), fileName: f.getName().split(".")[0].trim()});
    }
    var folders = folder.getFolders();
    while (folders.hasNext()) getFileIds(folders.next(), fileList, q);
    return fileList;
  };
  var sourceFiles = getFileIds(DriveApp.getFolderById(sourceFolderId), [], "mimeType='" + MimeType.MICROSOFT_EXCEL + "' or mimeType='" + MimeType.MICROSOFT_EXCEL_LEGACY + "'");
  var destinationFiles = getFileIds(DriveApp.getFolderById(destinationFolderId), [], "mimeType='" + MimeType.GOOGLE_SHEETS + "'");
  var createFiles = sourceFiles.filter(function(e) {return destinationFiles.every(function(f) {return f.fileName !== e.fileName});});
  var updateFiles = sourceFiles.reduce(function(ar, e) {
    var dst = destinationFiles.filter(function(f) {return f.fileName === e.fileName});
    if (dst.length > 0) {
      e.to = dst[0].id;
      ar.push(e);
    }
    return ar;
  }, []);
  if (createFiles.length > 0) createFiles.forEach(function(e) {Drive.Files.insert({mimeType: MimeType.GOOGLE_SHEETS, parents: [{id: destinationFolderId}], title: e.fileName}, DriveApp.getFileById(e.id))});
  if (updateFiles.length > 0) updateFiles.forEach(function(e) {Drive.Files.update({}, e.to, DriveApp.getFileById(e.id))});
  expectedpaymenttable(updateFiles[0].to,today);
 
}



function expectedpaymenttable(id,today){
 var sheet = SpreadsheetApp.openById(id);
 var ss   = sheet.getSheetByName("Sheet1");
 var startingRow;
 var endingRow;
 for(var b = 4; b <= ss.getLastRow(); b++) {
          var check = ss.getRange(b,1).getDisplayValue();
   
        var ending = today+' Total';
       
        if (check === today)
        {
          startingRow=b;
        }
        else if (check === ending)
        {
          endingRow=b;
        }
      }
  var range = ss.getRange(startingRow,1,endingRow-startingRow+1,ss.getLastColumn());
  var range2 = ss.getRange(4,1,1,ss.getLastColumn());
  var data = range2.getDisplayValues().concat(range.getDisplayValues());

var TABLEFORMAT = 'cellspacing="2" cellpadding="2" dir="ltr" border="1" style="width:100%;table-layout:fixed;font-size:10pt;font-family:arial,sans,sans-serif;border-collapse:collapse;border:1px solid #ccc;font-weight:normal;color:black;background-color:white;text-align:center;text-decoration:none;font-style:normal;'
var htmltable = '<table ' + TABLEFORMAT +' ">';

for (row = 0; row<data.length; row++){

htmltable += '<tr>';

for (col = 0 ;col<data[row].length; col++){
  if (data[row][col] === "" || 0) {htmltable += '<td>' + ' ' + '</td>';}
  else
    if (row === 0)  {
      htmltable += '<th>' + data[row][col] + '</th>';
    }

  else {htmltable += '<td>' + data[row][col] + '</td>';}
}

     htmltable += '</tr>';
}

     htmltable += '</table>';
     //Logger.log(data);
     //Logger.log(htmltable);
     var email = "REPLACE";
MailApp.sendEmail(email, 'Expected Payment '+today,'' ,{htmlBody: htmltable})
}


function dashboard(){
 var sourceFolderId = "1i7O50pLHwpIk5YFh8o-ECxhQQzajwUmd"; // Folder ID including source files.
  var destinationFolderId = "1wTHqTwHH5bnrJTqjPPfPEiq2bca_rHv8"; // Folder ID that the converted files are put.

  var getFileIds = function (folder, fileList, q) {
    var files = folder.searchFiles(q);
    while (files.hasNext()) {
      var f = files.next();
      fileList.push({id: f.getId(), fileName: f.getName().split(".")[0].trim()});
    }
    var folders = folder.getFolders();
    while (folders.hasNext()) getFileIds(folders.next(), fileList, q);
    return fileList;
  };
  var sourceFiles = getFileIds(DriveApp.getFolderById(sourceFolderId), [], "mimeType='" + MimeType.MICROSOFT_EXCEL + "' or mimeType='" + MimeType.MICROSOFT_EXCEL_LEGACY + "'");
  var destinationFiles = getFileIds(DriveApp.getFolderById(destinationFolderId), [], "mimeType='" + MimeType.GOOGLE_SHEETS + "'");
  var createFiles = sourceFiles.filter(function(e) {return destinationFiles.every(function(f) {return f.fileName !== e.fileName});});
  var updateFiles = sourceFiles.reduce(function(ar, e) {
    var dst = destinationFiles.filter(function(f) {return f.fileName === e.fileName});
    if (dst.length > 0) {
      e.to = dst[0].id;
      ar.push(e);
    }
    return ar;
  }, []);
  if (createFiles.length > 0) createFiles.forEach(function(e) {Drive.Files.insert({mimeType: MimeType.GOOGLE_SHEETS, parents: [{id: destinationFolderId}], title: e.fileName}, DriveApp.getFileById(e.id))});
  if (updateFiles.length > 0) updateFiles.forEach(function(e) {Drive.Files.update({}, e.to, DriveApp.getFileById(e.id))});
 
  var ss = SpreadsheetApp.openById("1fsTGPqMDzsELl4qlcpXIkHkNOL0rmVewt05LNIiZ5Ls");
  var statussheet = ss.getSheets();
  var BMOBalance;
  var REPLACEBANKALLBalance;
  var BMOQuantum;
  var BMOTime;
  var InterTime;
  var REPLACEBANK2Time;
  var REPLACEBANK2Quantum;
  var BalanceDate;
  var REPLACEBANK1Ending;
  var REPLACEBANK2Ending;
  var EuroBalance;
  var Interdate;
  var GBPBalance;
  var CollateralBalance;
  for(var i = 0; i <=5; i++) {
  var currentfile = updateFiles[i].to;
  var opensheet = SpreadsheetApp.openById(currentfile);
  var currentsheet = opensheet.getSheets();
  var currenttitle = currentsheet[0].getRange(1,1).getValues();
  var currenttitle2 = currentsheet[0].getRange(1,4).getValues();
 
  var BMOrow;
  var REPLACEBANK2row;
  var REPLACEBANK1EndingRow;
  var REPLACEBANK2EndingRow;
  var EuroRow;
  var GBPRow;
  var CollateralRow;
    if(currenttitle == ""&& currenttitle2 == ""){
      for(var b = 15; b <= currentsheet[0].getLastRow(); b++) {
          var check = currentsheet[0].getRange(b,2).getDisplayValue();
        //Logger.log(check);
        //var ending = today+' Total';
       
        if (check == "00021597213")
        {
          BMOrow=b;
        }
        else if (check == "5720081131")
        {
          REPLACEBANK2row=b;
        }
        else if (check == "051405515")
        {
          REPLACEBANK1EndingRow=b;
        }
        else if (check == "056073502")
        {
          REPLACEBANK2EndingRow=b;
        }
        else if (check == "8033045264")
        {
          EuroRow=b;
        }
        else if (check == "2081347274")
        {
          CollateralRow=b;
        }
        else if (check == "20325383386619")
        {
          GBPRow=b;
        }
       
      }
    currenttitle=currentsheet[0].getRange(2,21).getValues();
    BMOBalance=currentsheet[0].getRange(BMOrow,28).getValues();
    REPLACEBANKALLBalance=currentsheet[0].getRange(REPLACEBANK2row,28).getValues();
    REPLACEBANK1Ending=currentsheet[0].getRange(REPLACEBANK1EndingRow,28).getValues();
    REPLACEBANK2Ending= currentsheet[0].getRange(REPLACEBANK2EndingRow,28).getValues();
    EuroBalance= currentsheet[0].getRange(EuroRow,28).getValues();
    CollateralBalance= currentsheet[0].getRange(CollateralRow,28).getValues();
    try {
    GBPBalance= currentsheet[0].getRange(GBPRow,28).getValues();
    } catch (e) {
    Logger.log("GBP Bank Holiday")
    }
    BalanceDate=currentsheet[0].getRange(4,4).getValues();
    }
    else if (currenttitle == "PD Recon - Updated – Prior Day Match - AccountsSummary-All Accounts")
    {
      statussheet[0].getRange(2,8).setValue("TRUE");
      for(var b = 5; b <= currentsheet[0].getLastRow(); b++) {
          var check = currentsheet[0].getRange(b,1).getValues();
        if (check != "Detail Matched" && check != "Fully Reconciled / Marked")
        {
          statussheet[0].getRange(2,8).setValue("FALSE");
          break;
        }
      }
    }
    else if(currenttitle == "CashXplorer-Bank of Montreal - CD BMO")
    {
      BMOQuantum = currentsheet[0].getRange(currentsheet[0].getLastRow()-2,currentsheet[0].getLastColumn()).getValues();
      BMOTime= currentsheet[0].getRange(2,1).getValues();
      statussheet[0].getRange(3,2).setValue(BMOTime);
      statussheet[0].getRange(3,3).setValue(BMOQuantum);
    }
    else if(currenttitle == "CashXplorer-Intercompany Loans - Intercompany- US")
    {
      InterTime= currentsheet[0].getRange(2,1).getValues();
      statussheet[0].getRange(7,2).setValue(InterTime);
      for(var b = 5; b <= currentsheet[0].getLastRow(); b++) {
          var check = currentsheet[0].getRange(b,2).getDisplayValue();      
        if (check == "UL 10000-11000, UL 10000(REPLACEBANKALL) loans to 11000(REPLACEBANK1)")
        {
          statussheet[0].getRange(8,2).setValue(currentsheet[0].getRange(b,4).getDisplayValue());
          statussheet[0].getRange(8,3).setValue(currentsheet[0].getRange(b,5).getDisplayValue());
          statussheet[0].getRange(8,4).setValue(currentsheet[0].getRange(b,6).getDisplayValue());
          statussheet[0].getRange(8,5).setValue(currentsheet[0].getRange(b,7).getDisplayValue());
        }
        else if (check == "UL 10000-21100, UL 10000(REPLACEBANKALL) loans to 21100(REPLACEBANK2)")
        {
          statussheet[0].getRange(9,2).setValue(currentsheet[0].getRange(b,4).getDisplayValue());
          statussheet[0].getRange(9,3).setValue(currentsheet[0].getRange(b,5).getDisplayValue());
          statussheet[0].getRange(9,4).setValue(currentsheet[0].getRange(b,6).getDisplayValue());
          statussheet[0].getRange(9,5).setValue(currentsheet[0].getRange(b,7).getDisplayValue());
        }
        else if (check == "UL 21100-11000, UL 21100(REPLACEBANK2) loans to 11000(REPLACEBANK1)"){
          statussheet[0].getRange(10,2).setValue(currentsheet[0].getRange(b,4).getDisplayValue());
          statussheet[0].getRange(10,3).setValue(currentsheet[0].getRange(b,5).getDisplayValue());
          statussheet[0].getRange(10,4).setValue(currentsheet[0].getRange(b,6).getDisplayValue());
          statussheet[0].getRange(10,5).setValue(currentsheet[0].getRange(b,7).getDisplayValue());
        }}
    }
    else if(currenttitle == "CashXplorer-REPLACEBANK2 DDAs - CD Concentration Funding")
    {
      REPLACEBANK2Quantum = currentsheet[0].getRange(currentsheet[0].getLastRow()-2,currentsheet[0].getLastColumn()-6).getValues();
      REPLACEBANK2Time= currentsheet[0].getRange(2,1).getValues();
      statussheet[0].getRange(4,2).setValue(REPLACEBANK2Time);
      statussheet[0].getRange(4,3).setValue(REPLACEBANK2Quantum);
    }
    else if(currenttitle2="SecType")
    {
      Interdate=currentsheet[0].getRange(2,2).getValues();
    }
  Logger.log(currenttitle);
   }
  Logger.log(Interdate);
  statussheet[0].getRange(3,4).setValue(BMOBalance);
  statussheet[0].getRange(4,4).setValue(REPLACEBANKALLBalance);
  statussheet[0].getRange(5,2).setValue(BalanceDate);
  statussheet[0].getRange(6,2).setValue(Interdate);
  statussheet[0].getRange(12,4).setValue(REPLACEBANK1Ending);
  statussheet[0].getRange(13,4).setValue(REPLACEBANK2Ending);
  statussheet[0].getRange(14,4).setValue(EuroBalance);
  statussheet[0].getRange(16,4).setValue(GBPBalance);
  statussheet[0].getRange(15,4).setValue(CollateralBalance);
  const sheet = SpreadsheetApp.openById('1tF-wGPJ-Sg_Dw_9wf49YYs4AvXP8szpj6tCndLMOKUE').getSheets()[0];
  var yesterday=sheet.getRange(6,3).getDisplayValues();
  var suspensedate=sheet.getRange(6,1).getValues();
  statussheet[0].getRange(11,2).setValue(suspensedate);
  statussheet[0].getRange(11,8).setValue(yesterday);
}

function sendMail(id,today){
 var sh = SpreadsheetApp.openById(id);
 var data = sh.getDataRange().getDisplayValues();
  //var htmltable =[];

var TABLEFORMAT = 'cellspacing="2" cellpadding="2" dir="ltr" border="1" style="width:100%;table-layout:fixed;font-size:10pt;font-family:arial,sans,sans-serif;border-collapse:collapse;border:1px solid #ccc;font-weight:normal;color:black;background-color:white;text-align:center;text-decoration:none;font-style:normal;'
var htmltable = '<table ' + TABLEFORMAT +' ">';

for (row = 0; row<data.length; row++){

htmltable += '<tr>';

for (col = 0 ;col<data[row].length; col++){
  if (data[row][col] === "" || 0) {htmltable += '<td>' + ' ' + '</td>';}
  else
    if (row === 0)  {
      htmltable += '<th>' + data[row][col] + '</th>';
    }

  else {htmltable += '<td>' + data[row][col] + '</td>';}
}

     htmltable += '</tr>';
}

     htmltable += '</table>';
     Logger.log(data);
     Logger.log(htmltable);
     var email = "REPLACE";
MailApp.sendEmail(email, 'CAD Balances for '+today,'' ,{htmlBody: htmltable})
}

function BankRecon(){
var threads = GmailApp.search("QTM PROD-Quantum Bank Account Balance Report");// from today
var messages = threads[0].getMessages();
var message = messages[messages.length - 1];
var attachment = message.getAttachments()[0];
attachment.setContentType('application/vnd.ms-excel');
//var attachmentBlob = attachment.copyBlob();
//  var file = DriveApp.createFile(attachmentBlob);
//  Drive.Files.insert(
//          {
//            title: 'Bank Account Report',
//            mimeType: attachment.getContentType(),
//            parents: [{ id: '1i7O50pLHwpIk5YFh8o-ECxhQQzajwUmd' }]
//          },
//     attachment.copyBlob())
  Drive.Files.update({
    title: 'Bank Account Report', mimeType: attachment.getContentType()
  }, '1SadAL7Ulx5Ry7B1b8uiO_F2XDBNeNN0r', attachment.copyBlob());
 
  var threads = GmailApp.search("[External Sender] Interco Loan Report");// from today
var messages = threads[0].getMessages();
var message = messages[messages.length - 1];
var attachment = message.getAttachments()[0];
attachment.setContentType('application/vnd.ms-excel');
//var attachmentBlob = attachment.copyBlob();
//  var file = DriveApp.createFile(attachmentBlob);
//  Drive.Files.insert(
//          {
//            title: 'Bank Account Report',
//            mimeType: attachment.getContentType(),
//            parents: [{ id: '1i7O50pLHwpIk5YFh8o-ECxhQQzajwUmd' }]
//          },
//     attachment.copyBlob())
  Drive.Files.update({
    title: 'Interco Loan Report', mimeType: attachment.getContentType()
  }, '146NF9c2_QuK61nc2e6Ad0G_jmWb9V1Jo', attachment.copyBlob());
//}
 
            }

function getTextFromPDF() {
  var threads = GmailApp.search("QTM PROD-Quantum Suspense Report ");// from today
var messages = threads[0].getMessages();
var message = messages[messages.length - 1];
var attachment = message.getAttachments()[0];
attachment.setContentType('application/pdf');
  Drive.Files.update({
    title: 'Suspense Report', mimeType: attachment.getContentType()
  }, '1BdpstldzEVJHirbElIjiEbrto7iKrn-U', attachment.copyBlob());
  //var fileID='1BdpstldzEVJHirbElIjiEbrto7iKrn-U';
  var blob = DriveApp.getFileById('1BdpstldzEVJHirbElIjiEbrto7iKrn-U').getBlob()
  var resource = {
    title: blob.getName(),
    mimeType: blob.getContentType()
  };
  var options = {
  ocr: true,
    ocrLanguage: "en"
  };
   //Convert the pdf to a Google Doc with ocr.
  var file = Drive.Files.insert(resource, blob, options);
 
   //Get the texts from the newly created text.
  var doc = DocumentApp.openById(file.id);
  var text = doc.getBody().getText();
  Drive.Files.remove(doc.getId());
  var lines = text.split("\n")
 
  const sheet = SpreadsheetApp.openById('1tF-wGPJ-Sg_Dw_9wf49YYs4AvXP8szpj6tCndLMOKUE').getSheets()[0];
  //Logger.log(lines.length)
  for (i=0;i<lines.length;i++){
    sheet.getRange(i+1,1).setValue(lines[i]);
  }
  var today=sheet.getRange(6,2).getValues();
  if(lines.length>10 && today){
    suspensereport()
    return;
  }
}
  function suspensereport()
{
 
  var email = "REPLACE";
 
  var message = "Suspense Report Contains Item";

  var subject = "Please look at Suspense Report";
 
  MailApp.sendEmail(email, subject, message);

}
";
 
  var message = "Suspense Report Contains Item";

  var subject = "Please look at Suspense Report";
 
  MailApp.sendEmail(email, subject, message);

}
";
 
  var message = " \n"+dt+" \n"+" \n"+"REPLACEBANK1:  "+REPLACEBANK1+" \n"+" \n"+"REPLACEBANK2: "+REPLACEBANK2+" \n"+" \n"+"Total Fedline Balance: "+total;

  var subject = "Fedline Balance Update";
 
  MailApp.sendEmail(email, subject, message);

}
function fivepm(dt,REPLACEBANK2,REPLACEBANK1,total)
{

  var email = "REPLACE";
 
  var message = " \n"+dt+" \n"+" \n"+"REPLACEBANK1:  "+REPLACEBANK1+" \n"+" \n"+"REPLACEBANK2: "+REPLACEBANK2+" \n"+" \n"+"Total Fedline Balance: "+total;

  var subject = "5pm Fedline Balance";
 
  MailApp.sendEmail(email, subject, message);
}
function sixthirty(dt,REPLACEBANK2,REPLACEBANK1,total)
{
  const sheet = SpreadsheetApp.openById("1fsTGPqMDzsELl4qlcpXIkHkNOL0rmVewt05LNIiZ5Ls").getSheets()[0];
  var REPLACEBANK1numb = parseFloat(REPLACEBANK1.split("**")[REPLACEBANK1.split("**").length-1].split("$")[1].split(",").join(""))
  var REPLACEBANK2numb = parseFloat(REPLACEBANK2.split("**")[REPLACEBANK2.split("**").length-1].split("$")[1].split(",").join(""))
  sheet.getRange(12,3).setValue(REPLACEBANK1numb);
  sheet.getRange(13,3).setValue(REPLACEBANK2numb);
  var email = "REPLACE";
 
  var message = " \n"+dt+" \n"+" \n"+"REPLACEBANK1:  "+REPLACEBANK1+" \n"+" \n"+"REPLACEBANK2: "+REPLACEBANK2+" \n"+" \n"+"Total Fedline Balance: "+total;

  var subject = "Closing Fedline Balance";
 
  MailApp.sendEmail(email, subject, message);
}
function emailme(today)
{
  var REPLACEBANK1 = SpreadsheetApp.getActiveSheet().getRange(17,6).getDisplayValue();
  var REPLACEBANK2 = SpreadsheetApp.getActiveSheet().getRange(58,6).getDisplayValue();
  var email = "REPLACE";
 
  var message = "ACH, check and cash letter activity"+" \n"+"REPLACEBANK1:  "+REPLACEBANK1+" \n"+"REPLACEBANK2: "+REPLACEBANK2;

  var subject = "ACH & Cash Balances for "+today;
 
  MailApp.sendEmail(email, subject, message);

}
function emailstatus(status,date,msg,today)
{
 
  var ss = SpreadsheetApp.openById("1xz1hheHURpwkRojtyfbz02YVbpXmIh9pz_Ns8YfICdg");
  var sheets = ss.getSheets();
  var datestatus = sheets[0].getRange(2,2).setValue(today);
  var email = "REPLACE";
 
  var message = ""+date+" \n"+"Status:  "+status+" \n"+msg;

  var subject = "Fedline is "+status+" for "+date;
 
  MailApp.sendEmail(email, subject, message);
  hourly="99:99"
  EOD(hourly)
}
function extension(status,date,msg)
{
 
  var email = "REPLACE";
 
  var message = "Please check the current Status here: https://www.frbservices.org/app/status/serviceStatus.do";

  var subject = "Fedwire Funds Service has issued an EXTENSION";
 
  MailApp.sendEmail(email, subject, message);

}


function convert(today){
 
 var sourceFolderId = "1YUt_hTNClRUzhzrOOCPKas6_mkopTMRY"; // Folder ID including source files.
  var destinationFolderId = "1YUt_hTNClRUzhzrOOCPKas6_mkopTMRY"; // Folder ID that the converted files are put.

  var getFileIds = function (folder, fileList, q) {
    var files = folder.searchFiles(q);
    while (files.hasNext()) {
      var f = files.next();
      fileList.push({id: f.getId(), fileName: f.getName().split(".")[0].trim()});
    }
    var folders = folder.getFolders();
    while (folders.hasNext()) getFileIds(folders.next(), fileList, q);
    return fileList;
  };
  var sourceFiles = getFileIds(DriveApp.getFolderById(sourceFolderId), [], "mimeType='" + MimeType.MICROSOFT_EXCEL + "' or mimeType='" + MimeType.MICROSOFT_EXCEL_LEGACY + "'");
  var destinationFiles = getFileIds(DriveApp.getFolderById(destinationFolderId), [], "mimeType='" + MimeType.GOOGLE_SHEETS + "'");
  var createFiles = sourceFiles.filter(function(e) {return destinationFiles.every(function(f) {return f.fileName !== e.fileName});});
  var updateFiles = sourceFiles.reduce(function(ar, e) {
    var dst = destinationFiles.filter(function(f) {return f.fileName === e.fileName});
    if (dst.length > 0) {
      e.to = dst[0].id;
      ar.push(e);
    }
    return ar;
  }, []);
  if (createFiles.length > 0) createFiles.forEach(function(e) {Drive.Files.insert({mimeType: MimeType.GOOGLE_SHEETS, parents: [{id: destinationFolderId}], title: e.fileName}, DriveApp.getFileById(e.id))});
  if (updateFiles.length > 0) updateFiles.forEach(function(e) {Drive.Files.update({}, e.to, DriveApp.getFileById(e.id))});
  sendMail(updateFiles[0].to,today);
}
function expectedpayment(today){
 var sourceFolderId = "1s1neN4W8DHQY-8tBqlDj7n69-YHtgNCT"; // Folder ID including source files.
  var destinationFolderId = "1s1neN4W8DHQY-8tBqlDj7n69-YHtgNCT"; // Folder ID that the converted files are put.

  var getFileIds = function (folder, fileList, q) {
    var files = folder.searchFiles(q);
    while (files.hasNext()) {
      var f = files.next();
      fileList.push({id: f.getId(), fileName: f.getName().split(".")[0].trim()});
    }
    var folders = folder.getFolders();
    while (folders.hasNext()) getFileIds(folders.next(), fileList, q);
    return fileList;
  };
  var sourceFiles = getFileIds(DriveApp.getFolderById(sourceFolderId), [], "mimeType='" + MimeType.MICROSOFT_EXCEL + "' or mimeType='" + MimeType.MICROSOFT_EXCEL_LEGACY + "'");
  var destinationFiles = getFileIds(DriveApp.getFolderById(destinationFolderId), [], "mimeType='" + MimeType.GOOGLE_SHEETS + "'");
  var createFiles = sourceFiles.filter(function(e) {return destinationFiles.every(function(f) {return f.fileName !== e.fileName});});
  var updateFiles = sourceFiles.reduce(function(ar, e) {
    var dst = destinationFiles.filter(function(f) {return f.fileName === e.fileName});
    if (dst.length > 0) {
      e.to = dst[0].id;
      ar.push(e);
    }
    return ar;
  }, []);
  if (createFiles.length > 0) createFiles.forEach(function(e) {Drive.Files.insert({mimeType: MimeType.GOOGLE_SHEETS, parents: [{id: destinationFolderId}], title: e.fileName}, DriveApp.getFileById(e.id))});
  if (updateFiles.length > 0) updateFiles.forEach(function(e) {Drive.Files.update({}, e.to, DriveApp.getFileById(e.id))});
  expectedpaymenttable(updateFiles[0].to,today);
 
}



function expectedpaymenttable(id,today){
 var sheet = SpreadsheetApp.openById(id);
 var ss   = sheet.getSheetByName("Sheet1");
 var startingRow;
 var endingRow;
 for(var b = 4; b <= ss.getLastRow(); b++) {
          var check = ss.getRange(b,1).getDisplayValue();
   
        var ending = today+' Total';
       
        if (check === today)
        {
          startingRow=b;
        }
        else if (check === ending)
        {
          endingRow=b;
        }
      }
  var range = ss.getRange(startingRow,1,endingRow-startingRow+1,ss.getLastColumn());
  var range2 = ss.getRange(4,1,1,ss.getLastColumn());
  var data = range2.getDisplayValues().concat(range.getDisplayValues());

var TABLEFORMAT = 'cellspacing="2" cellpadding="2" dir="ltr" border="1" style="width:100%;table-layout:fixed;font-size:10pt;font-family:arial,sans,sans-serif;border-collapse:collapse;border:1px solid #ccc;font-weight:normal;color:black;background-color:white;text-align:center;text-decoration:none;font-style:normal;'
var htmltable = '<table ' + TABLEFORMAT +' ">';

for (row = 0; row<data.length; row++){

htmltable += '<tr>';

for (col = 0 ;col<data[row].length; col++){
  if (data[row][col] === "" || 0) {htmltable += '<td>' + ' ' + '</td>';}
  else
    if (row === 0)  {
      htmltable += '<th>' + data[row][col] + '</th>';
    }

  else {htmltable += '<td>' + data[row][col] + '</td>';}
}

     htmltable += '</tr>';
}

     htmltable += '</table>';
     //Logger.log(data);
     //Logger.log(htmltable);
     var email = "REPLACE";
MailApp.sendEmail(email, 'Expected Payment '+today,'' ,{htmlBody: htmltable})
}


function dashboard(){
 var sourceFolderId = "1i7O50pLHwpIk5YFh8o-ECxhQQzajwUmd"; // Folder ID including source files.
  var destinationFolderId = "1wTHqTwHH5bnrJTqjPPfPEiq2bca_rHv8"; // Folder ID that the converted files are put.

  var getFileIds = function (folder, fileList, q) {
    var files = folder.searchFiles(q);
    while (files.hasNext()) {
      var f = files.next();
      fileList.push({id: f.getId(), fileName: f.getName().split(".")[0].trim()});
    }
    var folders = folder.getFolders();
    while (folders.hasNext()) getFileIds(folders.next(), fileList, q);
    return fileList;
  };
  var sourceFiles = getFileIds(DriveApp.getFolderById(sourceFolderId), [], "mimeType='" + MimeType.MICROSOFT_EXCEL + "' or mimeType='" + MimeType.MICROSOFT_EXCEL_LEGACY + "'");
  var destinationFiles = getFileIds(DriveApp.getFolderById(destinationFolderId), [], "mimeType='" + MimeType.GOOGLE_SHEETS + "'");
  var createFiles = sourceFiles.filter(function(e) {return destinationFiles.every(function(f) {return f.fileName !== e.fileName});});
  var updateFiles = sourceFiles.reduce(function(ar, e) {
    var dst = destinationFiles.filter(function(f) {return f.fileName === e.fileName});
    if (dst.length > 0) {
      e.to = dst[0].id;
      ar.push(e);
    }
    return ar;
  }, []);
  if (createFiles.length > 0) createFiles.forEach(function(e) {Drive.Files.insert({mimeType: MimeType.GOOGLE_SHEETS, parents: [{id: destinationFolderId}], title: e.fileName}, DriveApp.getFileById(e.id))});
  if (updateFiles.length > 0) updateFiles.forEach(function(e) {Drive.Files.update({}, e.to, DriveApp.getFileById(e.id))});
 
  var ss = SpreadsheetApp.openById("1fsTGPqMDzsELl4qlcpXIkHkNOL0rmVewt05LNIiZ5Ls");
  var statussheet = ss.getSheets();
  var BMOBalance;
  var REPLACEBANKALLBalance;
  var BMOQuantum;
  var BMOTime;
  var InterTime;
  var REPLACEBANK2Time;
  var REPLACEBANK2Quantum;
  var BalanceDate;
  var REPLACEBANK1Ending;
  var REPLACEBANK2Ending;
  var EuroBalance;
  var Interdate;
  var GBPBalance;
  var CollateralBalance;
  for(var i = 0; i <=5; i++) {
  var currentfile = updateFiles[i].to;
  var opensheet = SpreadsheetApp.openById(currentfile);
  var currentsheet = opensheet.getSheets();
  var currenttitle = currentsheet[0].getRange(1,1).getValues();
  var currenttitle2 = currentsheet[0].getRange(1,4).getValues();
 
  var BMOrow;
  var REPLACEBANK2row;
  var REPLACEBANK1EndingRow;
  var REPLACEBANK2EndingRow;
  var EuroRow;
  var GBPRow;
  var CollateralRow;
    if(currenttitle == ""&& currenttitle2 == ""){
      for(var b = 15; b <= currentsheet[0].getLastRow(); b++) {
          var check = currentsheet[0].getRange(b,2).getDisplayValue();
        //Logger.log(check);
        //var ending = today+' Total';
       
        if (check == "00021597213")
        {
          BMOrow=b;
        }
        else if (check == "5720081131")
        {
          REPLACEBANK2row=b;
        }
        else if (check == "051405515")
        {
          REPLACEBANK1EndingRow=b;
        }
        else if (check == "056073502")
        {
          REPLACEBANK2EndingRow=b;
        }
        else if (check == "8033045264")
        {
          EuroRow=b;
        }
        else if (check == "2081347274")
        {
          CollateralRow=b;
        }
        else if (check == "20325383386619")
        {
          GBPRow=b;
        }
       
      }
    currenttitle=currentsheet[0].getRange(2,21).getValues();
    BMOBalance=currentsheet[0].getRange(BMOrow,28).getValues();
    REPLACEBANKALLBalance=currentsheet[0].getRange(REPLACEBANK2row,28).getValues();
    REPLACEBANK1Ending=currentsheet[0].getRange(REPLACEBANK1EndingRow,28).getValues();
    REPLACEBANK2Ending= currentsheet[0].getRange(REPLACEBANK2EndingRow,28).getValues();
    EuroBalance= currentsheet[0].getRange(EuroRow,28).getValues();
    CollateralBalance= currentsheet[0].getRange(CollateralRow,28).getValues();
    try {
    GBPBalance= currentsheet[0].getRange(GBPRow,28).getValues();
    } catch (e) {
    Logger.log("GBP Bank Holiday")
    }
    BalanceDate=currentsheet[0].getRange(4,4).getValues();
    }
    else if (currenttitle == "PD Recon - Updated – Prior Day Match - AccountsSummary-All Accounts")
    {
      statussheet[0].getRange(2,8).setValue("TRUE");
      for(var b = 5; b <= currentsheet[0].getLastRow(); b++) {
          var check = currentsheet[0].getRange(b,1).getValues();
        if (check != "Detail Matched" && check != "Fully Reconciled / Marked")
        {
          statussheet[0].getRange(2,8).setValue("FALSE");
          break;
        }
      }
    }
    else if(currenttitle == "CashXplorer-Bank of Montreal - CD BMO")
    {
      BMOQuantum = currentsheet[0].getRange(currentsheet[0].getLastRow()-2,currentsheet[0].getLastColumn()).getValues();
      BMOTime= currentsheet[0].getRange(2,1).getValues();
      statussheet[0].getRange(3,2).setValue(BMOTime);
      statussheet[0].getRange(3,3).setValue(BMOQuantum);
    }
    else if(currenttitle == "CashXplorer-Intercompany Loans - Intercompany- US")
    {
      InterTime= currentsheet[0].getRange(2,1).getValues();
      statussheet[0].getRange(7,2).setValue(InterTime);
      for(var b = 5; b <= currentsheet[0].getLastRow(); b++) {
          var check = currentsheet[0].getRange(b,2).getDisplayValue();      
        if (check == "UL 10000-11000, UL 10000(REPLACEBANKALL) loans to 11000(REPLACEBANK1)")
        {
          statussheet[0].getRange(8,2).setValue(currentsheet[0].getRange(b,4).getDisplayValue());
          statussheet[0].getRange(8,3).setValue(currentsheet[0].getRange(b,5).getDisplayValue());
          statussheet[0].getRange(8,4).setValue(currentsheet[0].getRange(b,6).getDisplayValue());
          statussheet[0].getRange(8,5).setValue(currentsheet[0].getRange(b,7).getDisplayValue());
        }
        else if (check == "UL 10000-21100, UL 10000(REPLACEBANKALL) loans to 21100(REPLACEBANK2)")
        {
          statussheet[0].getRange(9,2).setValue(currentsheet[0].getRange(b,4).getDisplayValue());
          statussheet[0].getRange(9,3).setValue(currentsheet[0].getRange(b,5).getDisplayValue());
          statussheet[0].getRange(9,4).setValue(currentsheet[0].getRange(b,6).getDisplayValue());
          statussheet[0].getRange(9,5).setValue(currentsheet[0].getRange(b,7).getDisplayValue());
        }
        else if (check == "UL 21100-11000, UL 21100(REPLACEBANK2) loans to 11000(REPLACEBANK1)"){
          statussheet[0].getRange(10,2).setValue(currentsheet[0].getRange(b,4).getDisplayValue());
          statussheet[0].getRange(10,3).setValue(currentsheet[0].getRange(b,5).getDisplayValue());
          statussheet[0].getRange(10,4).setValue(currentsheet[0].getRange(b,6).getDisplayValue());
          statussheet[0].getRange(10,5).setValue(currentsheet[0].getRange(b,7).getDisplayValue());
        }}
    }
    else if(currenttitle == "CashXplorer-REPLACEBANK2 DDAs - CD Concentration Funding")
    {
      REPLACEBANK2Quantum = currentsheet[0].getRange(currentsheet[0].getLastRow()-2,currentsheet[0].getLastColumn()-6).getValues();
      REPLACEBANK2Time= currentsheet[0].getRange(2,1).getValues();
      statussheet[0].getRange(4,2).setValue(REPLACEBANK2Time);
      statussheet[0].getRange(4,3).setValue(REPLACEBANK2Quantum);
    }
    else if(currenttitle2="SecType")
    {
      Interdate=currentsheet[0].getRange(2,2).getValues();
    }
  Logger.log(currenttitle);
   }
  Logger.log(Interdate);
  statussheet[0].getRange(3,4).setValue(BMOBalance);
  statussheet[0].getRange(4,4).setValue(REPLACEBANKALLBalance);
  statussheet[0].getRange(5,2).setValue(BalanceDate);
  statussheet[0].getRange(6,2).setValue(Interdate);
  statussheet[0].getRange(12,4).setValue(REPLACEBANK1Ending);
  statussheet[0].getRange(13,4).setValue(REPLACEBANK2Ending);
  statussheet[0].getRange(14,4).setValue(EuroBalance);
  statussheet[0].getRange(16,4).setValue(GBPBalance);
  statussheet[0].getRange(15,4).setValue(CollateralBalance);
  const sheet = SpreadsheetApp.openById('1tF-wGPJ-Sg_Dw_9wf49YYs4AvXP8szpj6tCndLMOKUE').getSheets()[0];
  var yesterday=sheet.getRange(6,3).getDisplayValues();
  var suspensedate=sheet.getRange(6,1).getValues();
  statussheet[0].getRange(11,2).setValue(suspensedate);
  statussheet[0].getRange(11,8).setValue(yesterday);
}

function sendMail(id,today){
 var sh = SpreadsheetApp.openById(id);
 var data = sh.getDataRange().getDisplayValues();
  //var htmltable =[];

var TABLEFORMAT = 'cellspacing="2" cellpadding="2" dir="ltr" border="1" style="width:100%;table-layout:fixed;font-size:10pt;font-family:arial,sans,sans-serif;border-collapse:collapse;border:1px solid #ccc;font-weight:normal;color:black;background-color:white;text-align:center;text-decoration:none;font-style:normal;'
var htmltable = '<table ' + TABLEFORMAT +' ">';

for (row = 0; row<data.length; row++){

htmltable += '<tr>';

for (col = 0 ;col<data[row].length; col++){
  if (data[row][col] === "" || 0) {htmltable += '<td>' + ' ' + '</td>';}
  else
    if (row === 0)  {
      htmltable += '<th>' + data[row][col] + '</th>';
    }

  else {htmltable += '<td>' + data[row][col] + '</td>';}
}

     htmltable += '</tr>';
}

     htmltable += '</table>';
     Logger.log(data);
     Logger.log(htmltable);
     var email = "REPLACE";
MailApp.sendEmail(email, 'CAD Balances for '+today,'' ,{htmlBody: htmltable})
}

function BankRecon(){
var threads = GmailApp.search("QTM PROD-Quantum Bank Account Balance Report");// from today
var messages = threads[0].getMessages();
var message = messages[messages.length - 1];
var attachment = message.getAttachments()[0];
attachment.setContentType('application/vnd.ms-excel');
//var attachmentBlob = attachment.copyBlob();
//  var file = DriveApp.createFile(attachmentBlob);
//  Drive.Files.insert(
//          {
//            title: 'Bank Account Report',
//            mimeType: attachment.getContentType(),
//            parents: [{ id: '1i7O50pLHwpIk5YFh8o-ECxhQQzajwUmd' }]
//          },
//     attachment.copyBlob())
  Drive.Files.update({
    title: 'Bank Account Report', mimeType: attachment.getContentType()
  }, '1SadAL7Ulx5Ry7B1b8uiO_F2XDBNeNN0r', attachment.copyBlob());
 
  var threads = GmailApp.search("[External Sender] Interco Loan Report");// from today
var messages = threads[0].getMessages();
var message = messages[messages.length - 1];
var attachment = message.getAttachments()[0];
attachment.setContentType('application/vnd.ms-excel');
//var attachmentBlob = attachment.copyBlob();
//  var file = DriveApp.createFile(attachmentBlob);
//  Drive.Files.insert(
//          {
//            title: 'Bank Account Report',
//            mimeType: attachment.getContentType(),
//            parents: [{ id: '1i7O50pLHwpIk5YFh8o-ECxhQQzajwUmd' }]
//          },
//     attachment.copyBlob())
  Drive.Files.update({
    title: 'Interco Loan Report', mimeType: attachment.getContentType()
  }, '146NF9c2_QuK61nc2e6Ad0G_jmWb9V1Jo', attachment.copyBlob());
//}
 
            }

function getTextFromPDF() {
  var threads = GmailApp.search("QTM PROD-Quantum Suspense Report ");// from today
var messages = threads[0].getMessages();
var message = messages[messages.length - 1];
var attachment = message.getAttachments()[0];
attachment.setContentType('application/pdf');
  Drive.Files.update({
    title: 'Suspense Report', mimeType: attachment.getContentType()
  }, '1BdpstldzEVJHirbElIjiEbrto7iKrn-U', attachment.copyBlob());
  //var fileID='1BdpstldzEVJHirbElIjiEbrto7iKrn-U';
  var blob = DriveApp.getFileById('1BdpstldzEVJHirbElIjiEbrto7iKrn-U').getBlob()
  var resource = {
    title: blob.getName(),
    mimeType: blob.getContentType()
  };
  var options = {
  ocr: true,
    ocrLanguage: "en"
  };
   //Convert the pdf to a Google Doc with ocr.
  var file = Drive.Files.insert(resource, blob, options);
 
   //Get the texts from the newly created text.
  var doc = DocumentApp.openById(file.id);
  var text = doc.getBody().getText();
  Drive.Files.remove(doc.getId());
  var lines = text.split("\n")
 
  const sheet = SpreadsheetApp.openById('1tF-wGPJ-Sg_Dw_9wf49YYs4AvXP8szpj6tCndLMOKUE').getSheets()[0];
  //Logger.log(lines.length)
  for (i=0;i<lines.length;i++){
    sheet.getRange(i+1,1).setValue(lines[i]);
  }
  var today=sheet.getRange(6,2).getValues();
  if(lines.length>10 && today){
    suspensereport()
    return;
  }
}
  function suspensereport()
{
 
  var email = "REPLACE";
 
  var message = "Suspense Report Contains Item";

  var subject = "Please look at Suspense Report";
 
  MailApp.sendEmail(email, subject, message);

}
