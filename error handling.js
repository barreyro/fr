var ERROR_LOG_SSID = "1EMHBPRAseHufnd1cQ3Ee8EfW8zVF7PPRn7GIBSXd6Hw";

function logErr_(err) {
  var stringErr = catchToString_(err);
  logErrInfo_(stringErr);
  return stringErr;
}

function catchToString_(err) {
  var errInfo = "Caught something:\n"; 
  for (var prop in err)  {  
    errInfo += "  property: "+ prop+ "\n    value: ["+ err[prop]+ "]\n"; 
  } 
  errInfo += "  toString(): " + " value: [" + err.toString() + "]"; 
  return errInfo;
}

function logErrInfo_(errInfo) {
  var ss = SpreadsheetApp.openById(ERROR_LOG_SSID);
  var sheet = ss.getSheets()[0];
  var date = new Date();
  var thisObj = {};
  thisObj.timestamp = date;
  thisObj.errorMessage = errInfo;
  sheet.getRange(sheet.getLastRow()+1, 1, 1, 2).setValues([[date, errInfo]]);
}

