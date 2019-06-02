/**
 * The event handler triggered when opening the spreadsheet.
 * @param {Event} e The onOpen event.
 */
function onOpen(e) {
  addRows();
  // Add a custom menu to the spreadsheet.
  SpreadsheetApp.getUi() // Or DocumentApp, SlidesApp, or FormApp.
      .createMenu('Custom Menu')
      .addItem('First item', 'menuItem1')
      .addToUi();
  
}

/**
 * Add rows to the sheet until there is a row for each date of the next year
 */
function addRows() {
  var activities = SpreadsheetApp.getActive().getSheetByName('Activities');
  // Make sure the Activities sheet is what is shown
  activities.activate();
  
  var nextDate = new Date();
  var endDate = new Date();
  endDate.setDate(endDate.getDate() + 9);
  //endDate.setFullYear(nextDate.getFullYear() + 1);

  // determine what is the current latest date
  var options = { weekday: 'short', year: 'numeric', month: 'short', day: 'numeric' };

  var lastRowIndex = activities.getLastRow();
  if (lastRowIndex > 1) {
    var cellData = activities.getSheetValues(lastRowIndex, 1, 1, 1);
    if (!(cellData[0][0] instanceof Date)) {
      var ui = SpreadsheetApp.getUi();
      ui.alert('Unexpected Input', 'There is unexpected input in the last row of this sheet. The date seems to have been modified. Please fix this and re-open the sheet', ui.ButtonSet.OK);
      return;
    }
    nextDate = cellData[0][0];
    nextDate.setDate(nextDate.getDate() + 1);
  }
  while (nextDate <= endDate) {
    activities.appendRow([nextDate]);
   
    nextDate.setDate(nextDate.getDate() + 1);
  }
  
  // format the column as expected
  activities.getRange('A2:A').setNumberFormat('DDD, MMM d, YYYY');
}