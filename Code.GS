/**
 * Get "home page", or a requested page.
 * Expects a 'page' parameter in querystring.
 *
 * @param {event} e Event passed to doGet, with querystring
 * @returns {String/html} Html to be served
 */
function doGet(e) {
  Logger.log( Utilities.jsonStringify(e) );
  if (!e.parameter.page) {
    // When no specific page requested, return "home page"
    return HtmlService.createTemplateFromFile('page1').evaluate();
  }
  // else, use page parameter to pick an html file from the script
  return HtmlService.createTemplateFromFile(e.parameter['page']).evaluate();
}
/**
 * Get the URL for the Google Apps Script running as a WebApp.
 */
function getScriptUrl() {
 var url = ScriptApp.getService().getUrl();
 return url;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

// ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


function AddRecord(serialNumber, partNumber, customer, courier, courierConsignmentNumber,soNumber,encryptionCheck, status, quantity, encryptionStatus) {
  var url = 'https://docs.google.com/spreadsheets/d/1jWgIbfKEQK_FEYwSXSL0PSm9mv-xn4uNW4wuvEqsypc/edit#gid=1425352652';
  var ss= SpreadsheetApp.openByUrl(url);
  var webAppSheet = ss.getSheetByName("CP_Data");
  var lastrow = webAppSheet.getLastRow()+1;
 
  const range = webAppSheet.getRange(lastrow,10);
  const imageURL = 'https://www.sanmina.com/wp-content/themes/FoundationPress-master/assets/images/sanmina-logo.png';
  const image  = SpreadsheetApp
  .newCellImage()
  .setSourceUrl(imageURL).setAltTextTitle('sanmina')
  .setAltTextDescription('sanmina logo')
  .build();

  webAppSheet.appendRow([new Date(),serialNumber, partNumber, customer, courier,courierConsignmentNumber,soNumber,encryptionCheck, status, image, quantity, encryptionStatus]);
  range.setValue(image)
  return "Console Point Database Record Added!"
}

//----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
// Reads Data From The Spreadsheet & Displays In The Form

function getInfo(partNumber) {
    const ss = SpreadsheetApp.openById("1jWgIbfKEQK_FEYwSXSL0PSm9mv-xn4uNW4wuvEqsypc");
    const ws = ss.getSheetByName("Full Part Info DB");
    const custIds = ws.getRange(2, 1, ws.getLastRow()-1, 1).getValues().map(r => r[0].toString());
    const posIndex = custIds.indexOf(partNumber.toString().toUpperCase());
    const rowNumber = posIndex === -1 ? 0 : posIndex + 2;
    console.log(rowNumber);
    if (rowNumber === 0){
        return {
                result:"Order Not Found",
                description:"OrderNotFound"
                }

    } 
    const customerInfo = ws.getRange(rowNumber, 1, 1, 8).getValues()[0]; //adjust 3rd number in rowNumber for the qty of columsn you want to read back - columns start at #1
    // const rdate = customerInfo[1].toISOString().split('T')[0];
    //console.log(rdate);
    return {
            result:"Order Found",
            description:customerInfo[2]
          }
          var res = "Hello";
}


// Reads Data From The Spreadsheet & Displays In The Form
function checkPartNum(partNum){

  var partNumber = partNum;
  //Open spreadsheet by url, point to sheet name 
  var url = 'https://docs.google.com/spreadsheets/d/1jWgIbfKEQK_FEYwSXSL0PSm9mv-xn4uNW4wuvEqsypc/edit#gid=1425352652';
  var ss = SpreadsheetApp.openByUrl(url);
  var ws = ss.getSheetByName("Full Part Info DB");

  //Get all rows of data, column 1 only (or where ever your primary key is) and split the data with a  comma
  var data = ws.getRange(1, 1, ws.getLastRow(), 1).getValues();
  var dJoin = data.join().split(",");
  Logger.log(dJoin);
  
  //Find the part number passed by the web for in the dataset.  Check if it returns a row or not...
  var myPnum = dJoin.indexOf(partNum);
  if (myPnum > -1){
    var dataRow = ws.getRange(myPnum + 1, 1, 1, ws.getLastColumn()).getValues();
    Logger.log(dataRow[0]);
    return dataRow[0];
    } else {
    var dataRow = [["Part number not recognised", "Part number not recognised","Part number not recognised"]];
    Logger.log(dataRow[0]);
    return dataRow[0];
  } 
 
}
 









// CODE FOR SN SEARCH FUNCTION //

function checkSN(findSN){


  //Open spreadsheet by url, point to sheet name 
  var url = 'https://docs.google.com/spreadsheets/d/1jWgIbfKEQK_FEYwSXSL0PSm9mv-xn4uNW4wuvEqsypc/edit#gid=1425352652';
  var ss = SpreadsheetApp.openByUrl(url);
  var ws = ss.getSheetByName("CP_Data");


  // <!-- //Get all rows of data, column 1 only (or where ever your primary key is) and split the data with a  comma -->
  // getRange(row, column, numRows, numColumns)
  // https://developers.google.com/apps-script/reference/spreadsheet/sheet#getrangerow,-column,-numrows,-numcolumns
  
  var data = ws.getRange(1, 2, ws.getLastRow(), 1).getValues();
  var dJoin = data.join().split(",");
  Logger.log(dJoin);
  

  // <!-- //Find the Serial number passed by the web for in the dataset.  Check if it returns a row or not... -->
  var mySnum = dJoin.indexOf(findSN);
  if (mySnum > -1){
    // var dataRow = ws.getRange(mySnum + 1, 1, 1, ws.getLastColumn()).getValues();
    var dataRow = ws.getRange(mySnum + 1, 2, 1, 7).getValues();
    Logger.log(dataRow[0]);
    return dataRow[0];
    } else {
    var dataRow = [["SN not recognised", "SN not recognised","SN not recognised"]];
    Logger.log(dataRow[0]);
    return dataRow[0];
  } 
 }


function getInfo(serialNumber) {
    const ss = SpreadsheetApp.openById("1jWgIbfKEQK_FEYwSXSL0PSm9mv-xn4uNW4wuvEqsypc");
    const ws = ss.getSheetByName("CP_Data");
    const custIds = ws.getRange(2, 2, ws.getLastRow()-1, 1).getValues().map(r => r[0].toString());
    const posIndex = custIds.indexOf(serialNumber.toString().toUpperCase());
    const rowNumber = posIndex === -1 ? 0 : posIndex + 2;
    console.log(rowNumber);
    if (rowNumber === 0){
        return {
                result:"Order Not Found",
                description:"OrderNotFound"
                }

    } 
    const customerInfo = ws.getRange(rowNumber, 1, 1, 16).getValues()[0]; //adjust 3rd number in rowNumber for the qty of columsn you want to read back - columns start at #1
    // const rdate = customerInfo[1].toISOString().split('T')[0];
    //console.log(rdate);
    return {
            result:"Order Found",
            description:customerInfo[2]
          }
          var res = "Hello";
}































