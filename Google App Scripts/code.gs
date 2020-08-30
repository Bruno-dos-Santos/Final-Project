function doGet(e){
  if (e.parameter.cmd == "doReview") { 
    if ( doReview(e) == true) {
      return HtmlService.createHtmlOutput('<h1>Done, the confirmation email was sent to your mailbox.</h1>')  
    } 
    else { 
      return HtmlService.createHtmlOutput('<h1>Something went wrong, it could had been reviewed before, ask finance team for more info! </h1>') 
    }
  }
  else 
  {
    return HtmlService.createHtmlOutputFromFile("index"); 
  }
}

// default variables to config the script
var inputSpreadsheetID = "1kwkVQS0LjhnhwwbpKFukHNCnqR1gIHIwxO_vSr-yhO4";
var decisionLink = 'https://script.google.com/macros/s/AKfycbyrwqSjbH69c9a7tTIZomnx3Ih8EBCdKztZSPxwAnYXmKfOn58/exec'; 
var folderID  = "1tNuNeiGrH-OW_1gDm-Fi7Sk-N3mmy_II";
var folderPDF = "16goOu5rroUPlKBL9QgRca0m5zOA0hij3";

var listStaff = {};


var statusIndex   = 4;
var htmlLM        = 49;
var htmlCFO       = 50;
var htmlCEO       = 51;


defaultEmailBody =
      '<p>Purchase Request\n'+ 'Date: {date} \n </p>' +
      '<p>Supplier Name: {supplierName} </p> ' + 
      '<p>Customer Relating To (If any): {CustomerRelating} </p>' +   
      '<p>Project Code: {ProjectCode} </p>' + 
      '<p>Reference Number: {referenceNumber} </p>'+  
      '<p>Department: {Department} </p>'+ 
      '<p>End User: {endUser} </p>' + 
      '<p>Reason for Purchase: </p>' +
      '{reason}\n' + 
      '<p>Currency: {currency} </p>'  + 
      '<p><b>Total: {total} </b></p>' +     
      '<p>Purchase Detail: ' + 
      '<ol><div>';
poCompanyDetails =  
  '<h2 style="text-align:center;color:#C02037;">{companyName}</h2> <br>' + 
  '<div>' + 
  '    <div> <b>' + 
  '        <div style="display:inline-block; width:40%;"> Billing Address</div>' + 
  '        <div style="display:inline-block; width:40%;"> Contact </div> </b>' + 
  '    </div>' + 
  '    <div>' + 
  '        <div style="display:inline-block; width:40%;"> {addressLine1} </div>' + 
  '        <div style="display:inline-block; width:40%;"> {contact} </div>' + 
  '    </div> ' + 
  '    <div>' + 
  '        <div style="display:inline-block; width:40%;"> {addressLine2} </div>' + 
  '        <div style="display:inline-block; width:40%;"> Tel: {tel}</div>' + 
  '    </div> ' + 
  '    <div>' + 
  '        <div style="display:inline-block; width:40%;"> {addressLine3} </div>' + 
  '        <div style="display:inline-block; width:40%;"> Fax: {fax}</div>' + 
  '    </div> ' + 
  '    <div>' + 
  '        <div style="display:inline-block; width:40%;"> {country} </div>' + 
  '        <div style="display:inline-block; width:40%;"> {emailFinance}</div>' + 
  '    </div> ' + 
  '</div>' + 
  '    <p><b> Please quote Purchase Order Number on All invoices and correspondence</b> ';


// this method update the default html above replacing the data from the row passed as parameter. 
function updateDefaultHtml(row)
{
  var result = defaultEmailBody;
  
  result = result.replace("{date}", getFormatedDate(row.date));
  result = result.replace("{supplierName}", row.supplier);
  result = result.replace("{CustomerRelating}", row.customerrelating);
  result = result.replace("{ProjectCode}", row.projectcode);
  result = result.replace("{referenceNumber}", row.referencenumber);
  result = result.replace("{Department}", row.department);
  result = result.replace("{endUser}", row.enduser);
  result = result.replace("{reason}", row.reason);
  result = result.replace("{currency}", row.currency);
  result = result.replace("{total}", formatFloat(row.total));
  
  for (items=1;items<=5;items++)
  {
    if (row["valitem"+items] != "") 
    {
      result = result + '' + 
        '<li> <div style="display:inline-block; text-align: left; width:25%;">' + row["descitem"+items] + '</div>' + 
          '<div style="display:inline-block; text-align: center; width:25%;">' + row["numitem"+items] + '</div>' +
            '<div style="display:inline-block; text-align: right; width:25%;">' + formatFloat(row["valitem"+items]) + '</div>' +  
              '<div style="display:inline-block; text-align: right; width:25%;">' + formatFloat(row["subtotalitem"+items]) + '</div></li>';
    } 
  }
  result = result + '</ol></div>';
  
  return result;
}


//this method update the company's details from the "Config" sheet from Google Docs.

function updateCompanyDetails()
{  
  var ss = SpreadsheetApp.openById(inputSpreadsheetID);
  var ws = ss.getSheetByName("Config");  
  poCompanyDetails = poCompanyDetails.replace("{companyName}", ws.getRange(2, 1).getValue());
  poCompanyDetails = poCompanyDetails.replace("{addressLine1}", ws.getRange(2, 2).getValue());
  poCompanyDetails = poCompanyDetails.replace("{addressLine2}", ws.getRange(2, 3).getValue());
  poCompanyDetails = poCompanyDetails.replace("{addressLine3}", ws.getRange(2, 4).getValue());
  poCompanyDetails = poCompanyDetails.replace("{contact}", ws.getRange(2, 8).getValue());
  poCompanyDetails = poCompanyDetails.replace("{tel}", ws.getRange(2, 6).getValue());  
  poCompanyDetails = poCompanyDetails.replace("{fax}", ws.getRange(2, 7).getValue());
  poCompanyDetails = poCompanyDetails.replace("{country}", ws.getRange(2, 5).getValue());  
  var emailFinance = ws.getRange(2, 9).getValue();
  poCompanyDetails = poCompanyDetails.replace("{emailFinance}", emailFinance);   
  return emailFinance;
}


// this method converts html to pdf, as part as the final step from the system.
function htmlToPDF(html, email, fileName, subject) {
  var htmlBody = "<p>Please find attached the purchase order.";
  
  var blob = Utilities.newBlob(html, "text/html", fileName + ".html");
  var pdf = blob.getAs("application/pdf");

  DriveApp.getFolderById(folderPDF).createFile(pdf).setName(fileName + ".pdf");

  MailApp.sendEmail(email, subject, "",
     {htmlBody: htmlBody, attachments: pdf});
}


//this is the main function, it's from here that the script would know what to do to each register inside the Google Docs
// full explanation is written over the final report doc.
function onReportOrApprovalSubmit() 
{
   function doAddLinks(link1, link2)
   {
     function addLink(urlLink, num)
     {
        if (urlLink != "") 
        {
          html = html + " <a href=" + urlLink + ">File Attached " + num + " </a> <br>";      
        }
     }
       addLink(row.attachment1, 1);    
       addLink(row.attachment2, 2);          
   } // end of doAddLinks
  
   function doSendEmail(email, subject, finalHtml,status)
   {
     sheet.getRange(startRow + i, 4).setValue(status);
     // keep the html, so the system can resend emails everyday.     
     if (status == "LM_SENT") 
     {
       sheet.getRange(startRow + i, htmlLM).setValue(finalHtml);      
     } 
     else if (status == "CFO_SENT") 
     {
       sheet.getRange(startRow + i, htmlCFO).setValue(finalHtml);            
     } 
     else if (status == "CEO_SENT") 
     {
       sheet.getRange(startRow + i, htmlCEO).setValue(finalHtml);            
     } 
     
     if (status == "PO_CREATED") 
     {
       var fileName = '{id}-{initials}';
       fileName = fileName.replace('{initials}', getInitials(staffName));                              
       fileName = fileName.replace('{id}',subject);        
      
       htmlToPDF(finalHtml, email, fileName, subject);           
     }
     else
     {
       MailApp.sendEmail(email, subject, "", {htmlBody:finalHtml});         
     }
    
     
     SpreadsheetApp.flush();
   } // end of doSendEmail
  
  function addReviewLink(level, replayEmail, supplierName, amt, html) 
  { 
    result = html;
    supplierName = supplierName.replace(/\s+/g, "%20")
    var urlApp = decisionLink +'?cmd=doReview&level=' + level + '&id=' + (i + startRow) + "&supplier="+ supplierName + "&amt=" + amt;
    var approve = urlApp + '&approval=true'+'&reply=' + replayEmail; 
    var reject = urlApp + '&approval=false'+'&reply=' + replayEmail;
    
    result =
      '<a target="_self" href='+  approve +'>Approve</a><br />'+
      '<a target="_self" href=' + reject +">Reject</a><br />" +
          result;
    return result;
  }
  
  
  function doPO()
  {
    // new requests
    var poNumber = getNextPO();   
    doAddLinks();
    poNumber = "PO-" + poNumber 
    var emailAddress = row.userid + "," + emailFinance;  // add email accounts          
    html = "<p><b>PO NUMBER: " + poNumber + "</p></b>" + poCompanyDetails +  html;              
    html = "<body>" + html + '</body>';
    setPONumber(poNumber);
    doSendEmail(emailAddress, poNumber, html, "PO_CREATED");      
  }
  
  function setPONumber(poNumber)
  {
    sheet.getRange(startRow + i, 2).setValue("Completed");
    sheet.getRange(startRow + i, 48).setValue(poNumber);
    // since the po has been created, we clear up the html to save space.
    sheet.getRange(startRow + i, htmlLM).setValue("");      
    sheet.getRange(startRow + i, htmlCFO).setValue("");            
    sheet.getRange(startRow + i, htmlCEO).setValue("");            
    SpreadsheetApp.flush();
  }
  
    
  //get staffs
  getStaff();
  var emailFinance = updateCompanyDetails();
  // read the data from the spreadsheet and begin the function itself. 
  var ss = SpreadsheetApp.openById(inputSpreadsheetID);
  var sheet = ss.getSheetByName("Data");
  var summary = ss.getSheets()[1];
  var startRow = 2;
  var data = getRowsData(sheet, sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn()));
  //walk through the PR opened.
  for (var i = 0; i < data.length; ++i) 
  {
    var row = data[i];
    var prID = i + startRow;
    if (row.status == "Created") 
    {
      var html = updateDefaultHtml(row); 
      var staffName = getStaffNameByEmail(row.userid);
      // subject 
      var subject = 'Review PR (date: {date}, created by {initials}, supplier: {supplier}, amt: {amt})';
      subject = subject.replace('{initials}', getInitials(staffName));                              
      subject = subject.replace('{date}', getFormatedDate(row.date));
      subject = subject.replace('{supplier}', row.supplier);
      subject = subject.replace('{amt}', formatFloat(row.total));
      switch (row.stage) 
      {
          case "EMAIL_SENT":
              var emailAddress = getStaffEmail(row.lmapprovalname); // email address                
              doAddLinks();              
              html = addReviewLink(1,emailAddress + "," + row.userid, row.supplier, row.total, html);              
              html = "<p> Hi " + row.lmapprovalname + ", please review the follow purchase request opened by " + staffName + "</p>" + html;
              html = "<body>" + html + '</body>'              
              doSendEmail(emailAddress, subject, html, "LM_SENT");            
            break;
          case "LM_APPROVED":     
              var emailAddress = getStaffEmail(row.cfoapprovalname); // email address                
              doAddLinks();              
              html = addReviewLink(2,emailAddress + "," + row.userid, row.supplier, row.total, html);                          
              html = "<p> Hi " + row.cfoapprovalname  + ", please review the follow purchase request approved by " + row.lmapprovalname + " and opened by " + staffName + "</p>"  + html;
              html = "<body>" + html + '</body>'              
              doSendEmail(emailAddress, subject, html, "CFO_SENT");              
            break;          
          case "CFO_APPROVED":
              var emailAddress = getStaffEmail(row.ceoapprovalname); // email address                
              doAddLinks();              
              html = addReviewLink(3,emailAddress + "," + row.userid, row.supplier, row.total, html);              
              html = "<p> Hi " + row.ceoapprovalname  + ", please review the follow purchase request approved by " + row.cfoapprovalname + " and opened by " + staffName + "</p>"  + html;
              html = "<body>" + html + '</body>'              
              doSendEmail(emailAddress, subject, html, "CEO_SENT");             
            break; 
          case "CEO_APPROVED":
            doPO()
            break; 
          case "LM_SENT":
            break;        
          case "CFO_SENT":          
            break;
          case "CFO_REJECTED":
            break; 
          case "CEO_SENT":
            break;
          case "CEO_REJECTED":
            break;      
          case "LM_REJECTED":
            break;         
          case "PO_CREATED":              
            break;                              
          default:
            // new request. 
            doAddLinks();
            var emailAddress = row.userid; // email address 
            html = "<body>" + html + '</body>';
            doSendEmail(emailAddress, subject, html, "EMAIL_SENT");  
          
      } // end of switch 
    }
  } //end of the for. 
}

// this method processes the review from the link inside the email, either approves or reject the data.
function doReview(e) {
  if (e.parameter.cmd == "doReview") {
    var ss = SpreadsheetApp.openById(inputSpreadsheetID);
    var ws = ss.getSheetByName("Data");
    id  = e.parameter.id;
    email = e.parameter.reply;
    row = 6;
    if (e.parameter.approval == "true") {
      str = "APPROVED"
    }
    else {
      str = "REJECTED"
    }
    status = "LM";
    
    if (e.parameter.level == "2") {
      status = "CFO";
      row    = 9;
      email = email;
    } else if  (e.parameter.level == "3") {
      status = "CEO";
      row    = 12; 
      email = email;
    }
    date = ws.getRange(id, row+1).getValue();
    if (date == "") {
      
      MailApp.sendEmail(email, "Your PR from " + e.parameter.supplier + " amt: " + e.parameter.amt + " was: " + str + " by " + status, "EOM");
      
      if (str == "REJECTED") {
        ws.getRange(id, 2).setValue("Completed");              
        
        // since the po has been rejected, we clear up the html to save space.
        ws.getRange(id, htmlLM).setValue("");      
        ws.getRange(id, htmlCFO).setValue("");            
        ws.getRange(id, htmlCEO).setValue("");            
      }
      status = status + "_" + str
      ws.getRange(id, 4).setValue(status);
      ws.getRange(id, row).setValue(str + " by the email!");
      ws.getRange(id, row+1).setValue(new Date());
      SpreadsheetApp.flush();
      return true
    }
    else{
      return false
    }
  }
}


// this method goes to the config sheet, gets the last PO number there and adds 1.
function getNextPO(){
  var ss = SpreadsheetApp.openById(inputSpreadsheetID);
  var ws = ss.getSheetByName("Config");  
  var lastPO = ws.getRange(2, 10).getValue();
  var nextPO = lastPO + 1;  
  ws.getRange(2,10).setValue(nextPO);
  SpreadsheetApp.flush();  
  return nextPO;
};


//this method reads the staff sheet to get the details from the staff, like email, full name and returns as a list. 
function getStaff(){
  listStaff = {};
  var ss = SpreadsheetApp.openById(inputSpreadsheetID);
  var ws = ss.getSheetByName("Staff");
  var dataRange = ws.getRange(2, 1, ws.getLastRow()-1, ws.getLastColumn());
  var data = dataRange.getValues();
 
  for (var i = 0; i < data.length; ++i){
    row = data[i];
    listStaff[row[2]] = [row[4], row[0]]; 
    listStaff[row[0]] = row[2];
  }
  Logger.log(listStaff)
}


// this method inputs the data from the frontend on the Google Spreadsheet, using the "Data" sheet as database.
function startSubmit(userInputData, file1, file2){
  getStaff();
  var ss = SpreadsheetApp.openById(inputSpreadsheetID);
  var ws = ss.getSheetByName("Data");
  var userEmail = getUserDetails();
  
  userInputData.unshift(new Date(), "Created", userEmail ,"",getStaffLmByEmail(userEmail),"","","CFO Joe","","","CEO John","","");
  
  var fileURL1 = "";
  var fileURL2 = "";
  if (file1 != undefined) {
    fileURL1 = uploadFileToDrive(file1[1], file1[0])
  } 
  
  if (file2 != undefined) {
    fileURL2 = uploadFileToDrive(file2[1], file2[0])    
  }  
  userInputData.push(fileURL1, fileURL2, "", "", "", "")  
  ws.appendRow(userInputData);
  
  
  onReportOrApprovalSubmit();
}



// this method gets the name from the staff list and converts that as initial using regEx 
function getInitials(name){
  var initials = name.match(/\b\w/g) || [];
  initials =  initials.join('');
  return initials
}

// this method gets the user details looged on google and returns its email.
function getUserDetails(){
  return Session.getActiveUser().getEmail(); 
}

//this method gets the staff's email from the staff list 
function getStaffEmail(name){
  return listStaff[name]
}

//this method gets the staff's LM from the staff list 
function getStaffLmByEmail(email){
  return listStaff[email][0]
}

//this method gets the staff's name from the staff list 
function getStaffNameByEmail(email){
  return listStaff[email][1]
}

//this method formats the date as readale format. 
function getFormatedDate(date) 
{
  return  Utilities.formatDate(new Date(date), "GMT", "dd/MM/yyyy")
}

// this method gets the row from the sheet assigned as parameter.
function getRowsData(sheet, range, columnHeadersRowIndex) {
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
//  Browser.msgBox(headers.toSource());
  return getObjects(range.getValues(), normalizeHeaders(headers));
//  return getObjects(range.getRowIndex);
}


// this method gets the object and convert them as key-value pairs. 
function getObjects(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      /* if (isCellEmpty(cellData)) {
        continue;
      } */ 
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}


// Normalize a string, by removing all non-alphanumeric characters and using mixed case
// to separate words.
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader(headers[i]);
    if (key.length > 0) {
      Logger.log("HDR:"+key);
      keys.push(key);
    }
  }
  return keys;
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit(char) {
  return char >= '0' && char <= '9';
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string

function isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}


// returns true if the value is empty
function isEmpty(value) {
  return  value == null || value == undefined || value == "";
}

//formats a float to present it over the html, aligned to the right. 
function formatFloat(num)
{
  numFloat = parseFloat(num) || 0
  return numFloat.toFixed(2).replace(/(\d)(?=(\d{3})+\.)/g, '$1,'); 
}

// uploads the base64Data to google drive and returns its pacth. 
function uploadFileToDrive(base64Data, fileName) {
  try{
    var splitBase = base64Data.split(','),
        type = splitBase[0].split(';')[0].replace('data:','');
    var byteCharacters = Utilities.base64Decode(splitBase[1]);
    var ss = Utilities.newBlob(byteCharacters, type);
    ss.setName(fileName);
    var file = DriveApp.getFolderById(folderID).createFile(ss);
    return file.getUrl();
  } catch(e){
    return 'createFile Error: ' + e.toString();
 }
}