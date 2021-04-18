/*********************************************************************************************************
*
* Instructions
* 1. Create a new Google Sheet. 
* 2. Open Google Apps Script.
* 3. At top, click Resources -> Libraries -> add this Script ID: 1Mc8BthYthXx6CoIz90-JiSzSafVnT6U3t0z_W3hLTAX5ek4w0G_EIrNw
and select Version 8 (or latest version). Save.
* 4. Delete all text in the scripting window and paste all this code.
* 5. Run onOpen(). Accept the permissions.
* 6. Then run "Return Home Depot Price from URL" from the spreadsheet.
*
*********************************************************************************************************/

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Functions')
  .addItem('Return Home Depot Price from URL', 'parseObject')
  .addToUi();
}

/*********************************************************************************************************
*
* Scrape web content.
* 
* @param {String} query The search URL to parse
*
* @return {String} Desired web page content.
*
* References
* https://www.reddit.com/r/spreadsheets/comments/mnw01b/how_can_i_automatically_pull_material_values_from/
* https://www.homedepot.com/
* https://www.kutil.org/2016/01/easy-data-scrapping-with-google-apps.html
*
*********************************************************************************************************/

function getData(query) {
  var fromText = '},"offers":';
  var toText = ',"review":';

  var content = UrlFetchApp.fetch(query).getContentText();
  
  // For debugging, download source text of URL to Google Drive since it's too much text for console log
  //  DriveApp.createFile("Website Data.txt", content);
  
  var scraped = Parser
  .data(content)
  .setLog()
  .from(fromText)
  .to(toText)
  .build();
  console.log(scraped);
  return scraped;
}

/*********************************************************************************************************
*
* Print scraped web content to Google Sheet.
* 
*********************************************************************************************************/

function parseObject(){
  
  //  Declare variables
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = spreadsheet.getActiveSheet();
  var currentCell = currentSheet.getCurrentCell();
  var currentCellValue = currentCell.getDisplayValue();  
  var searchArray = [{'query': currentCellValue}];
  
  //  Return website data, convert to Object
  var responseText = getData(searchArray[0].query);
  var responseTextJSON = JSON.parse(responseText);  
  
  //  Select object key with data in it
  searchArray[0].returnKey = responseTextJSON.price;
  
//  Print to cell to the right of current cell, uncomment to use
 currentSheet.getRange(currentCell.getRow(), currentCell.getColumn() + 1).setValue(searchArray[0].returnKey);
}
