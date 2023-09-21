function extractAccessReviewInfo() {
  const spreadsheetId =
    PropertiesService.getScriptProperties().getProperty('spreadsheetId');
  const userListSheetNameList = [
    '対象のグループアドレスリスト',
    '協力先リスト',
  ];
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const groupUsersListSheet = addWkSheetAndSetFormula_(spreadsheet);
  const userListSheet = createSheet_(spreadsheet, '入室可能者リスト');
  const tempFormulaText = userListSheetNameList
    .map(sheetName => `'${sheetName}'!A:A`)
    .join('; ');
  const userListFormula = `=QUERY({${tempFormulaText}}, "SELECT * WHERE Col1 IS NOT NULL")`;
  userListSheet.getRange('A1').setValue(userListFormula);
  const groupUsersListColumnRowIndex = 5;
  SpreadsheetApp.flush();
  const groupUsers = groupUsersListSheet.getDataRange().getValues();
  const targetGroupAddresses = JSON.parse(
    PropertiesService.getScriptProperties().getProperty('targetGroupAddresses')
  );
  console.log(groupUsersListSheet.getDataRange().getA1Notation());
  const groupUsersTargetColumnIndexList = findCommonElementIndices_(
    groupUsers[groupUsersListColumnRowIndex],
    targetGroupAddresses
  );
  const targetGroupUsers = groupUsers
    .map(value => groupUsersTargetColumnIndexList.map(index => value[index]))
    .filter((_, idx) => idx > groupUsersListColumnRowIndex);
  const targetUsers = flattenAndRemoveDuplicates_(targetGroupUsers);
  const targetUsersSheet = createSheet_(spreadsheet, userListSheetNameList[0]);
  targetUsersSheet.getRange(1, 1, targetUsers.length, 1).setValues(targetUsers);
  groupUsersListSheet.hideSheet();
  targetUsersSheet.hideSheet();
  userListSheet.hideSheet();
  const cooperationPartnersListSheet = createSheet_(
    spreadsheet,
    userListSheetNameList[1]
  );
  cooperationPartnersListSheet.showSheet();
  const printSheet = spreadsheet.getSheetByName('ISO印刷用');
  printSheet
    .getRange('H8')
    .setValue(
      PropertiesService.getScriptProperties().getProperty('formulaForReview')
    );
  printSheet.showColumns(8);
  SpreadsheetApp.flush();
}

function findCommonElementIndices_(arrayA, arrayB) {
  return arrayA
    .map((value, idx) => (arrayB.includes(value) ? idx : null))
    .filter(x => x !== null);
}

function flattenAndRemoveDuplicates_(arr) {
  // Flatten the two-dimensional array and remove empty values
  const flatArray = arr.flat().filter(value => value !== '');

  // Create a Set to remove duplicates and convert back to an array
  const uniqueArray = [...new Set(flatArray)].map(value => [value]);

  return uniqueArray;
}

function checkDateValue_(dateValue) {
  // Get today's date
  const today = new Date();
  today.setHours(0, 0, 0, 0); // Set hours, minutes, seconds, and milliseconds to zero
  return !dateValue || (dateValue instanceof Date && dateValue > today);
}

function addWkSheetAndSetFormula_(spreadsheet) {
  const formula =
    PropertiesService.getScriptProperties().getProperty('formula');
  const sheetName = 'wk';
  const sheet = createSheet_(spreadsheet, sheetName);
  sheet.getRange('A1').setFormula(formula);
  SpreadsheetApp.flush();
  Utilities.sleep(1000);
  sheet.getDataRange().setValues(sheet.getDataRange().getValues());
  return sheet;
}

function createSheet_(spreadsheet, newSheetName) {
  const sheet = spreadsheet.getSheetByName(newSheetName);
  if (sheet) {
    sheet.clear();
    return sheet;
  }
  const newSheet = spreadsheet.insertSheet(newSheetName);
  newSheet.setName(newSheetName);
  return newSheet;
}
