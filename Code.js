var scriptTitle = "formRanger Add-on V1.0 (9/8/14)";
var scriptName =  "formrangerAddOn";
var scriptTrackingId = "UA-48800213-3"; 

var DIALOG_TIMEOUT_SECONDS = 60;
var APPROVED_TYPES = ["CHECKBOX","LIST","MULTIPLE_CHOICE", "GRID"];

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
  .getContent();
}

function onInstall(e) {
 onOpen(e);
 formRangerUi();
}

function onOpen(e) {
  var menuItems = [];
  var menu = FormApp.getUi().createAddonMenu().addItem('Start', 'formRangerUi');
  menu.addToUi();
}


function formRangerUi() {
  setSid_();
  try {
    DriveApp.getRootFolder();
  } catch(err) {
    FormApp.getUi().alert("Oops! It looks like your domain administrator does not allow users to install 3rd party Apps that rely on the Google Drive API.");
    return;
  }
  var app = HtmlService.createTemplateFromFile("Range Sidebar").evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setTitle("formRanger"); //.setSandboxMode(HtmlService.SandboxMode.IFRAME)
  FormApp.getUi().showSidebar(app);
}


function reloadPanel() {
  try {
    populateAllFormQuestions();
    var triggerStateObj = {};
    triggerStateObj.triggerState = getTriggerState();
    triggerStateObj.timeTriggerState = getTimeTriggerState();
    return triggerStateObj;
  } catch(err) {
    err = logErr_(err);
    throw err;
  }
}


function showNewRangePicker(id) {
  var dialogId = id;
  var template = HtmlService.createTemplateFromFile("rangeMaker");
  template.dialogId = dialogId;
  template.ssId = '';
  template.sheetObjs = '';
  template.selectedSheet = '';
  template.selectedHeader = '';
  template.rangeName = '';
  template.rangeId = '';
  template.mode = "new";
  var page = template.evaluate()
  .setTitle("formRanger")
  .setTitle('Dialog')
  .setWidth(600).setHeight(450)
  .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  FormApp.getUi().showModalDialog(page, "Name and attach values list"); 
  checkIn(dialogId);
  return dialogId;
}


function showRangeEditor(questionId, rangeId) {
  var dialogId = questionId;
  var range = getRangeById(rangeId);
  var ss = SpreadsheetApp.openById(range.ssId);
  var sheet = NVSL.getSheetById(range.sheetId, ss);
  var template = HtmlService.createTemplateFromFile("rangeMaker");
  template.dialogId = dialogId;
  template.ssId = range.ssId;
  template.sheetObjs = JSON.stringify(getSheetObjs(range.ssId));
  template.selectedSheet = JSON.stringify(getSheetObj(range.ssId, range.sheetId));
  template.selectedHeader = range.colHeading;
  template.rangeName = range.name;
  template.rangeId = rangeId;
  template.mode = "edit";
  var page = template.evaluate()
  .setTitle("formRanger")
  .setTitle('Dialog')
  .setWidth(600).setHeight(450)
  .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  FormApp.getUi().showModalDialog(page, "Edit values list");
  checkIn(dialogId);
  return dialogId;
}


function getUiGlobal() {
  try {
    var ui_global = {};
    ui_global.activeFormQuestions = getActiveFormQuestions(); 
    ui_global.namedRanges = getNamedRanges();
    var doctopusRosters = buildRosterOptions();
    for (var i=0; i<doctopusRosters.length; i++) {
      ui_global.namedRanges.push(doctopusRosters[i]);
    }
    ui_global.questions = getActiveFormQuestions();
    ui_global.questionJobs = getQuestionJobs();
    ui_global.assignedQuestions = [];
    var eligibleQuestionIds = getEligibleQuestionIds();
    for (var i=0; i<ui_global.questionJobs.length; i++) {
      if (eligibleQuestionIds.indexOf(ui_global.questionJobs[i].questionId) === -1) {
         //garbage collection
         deleteQuestionJob(ui_global.questionJobs[i].questionId); 
      } else {
        ui_global.assignedQuestions.push(ui_global.questionJobs[i].questionId);
      }
    }
    ui_global = JSON.stringify(ui_global);
    return ui_global;
  } catch(err) {
    err = logErr_(err);
    throw err;
  }
}


function getActiveFormQuestions() {
  var form = FormApp.getActiveForm();
  try {
    var destId = form.getDestinationId();
  } catch(err) {
  }
  
  if (destId) {
    var destSS = SpreadsheetApp.openById(destId);
    var destTitle = destSS.getName();
    var destUrl = destSS.getUrl();
  }
  
  var items = form.getItems();
  var questionArray = [];
  var documentProperties = PropertiesService.getDocumentProperties().getProperties();
  for (var i=0; i<items.length; i++) {
    var thisQuestion = {};
    thisQuestion.title = items[i].getTitle();
    thisQuestion.type = items[i].getType().toString();
    thisQuestion.id = items[i].getId().toString();
    if (destId) {
      thisQuestion.destId = destId;
      thisQuestion.destTitle = destTitle;
      thisQuestion.destUrl = destUrl;
    }
    if (documentProperties.assignedRanges) {
      var assignedRanges = documentProperties.assignedRanges;
      if (assignedRanges['Q_'+thisQuestion.id]) {
        thisQuestion.assignedRange = JSON.parse(assignedRanges['Q_'+thisQuestion.id]);
      }
    }
    questionArray.push(thisQuestion);
  }
  return questionArray;
}



function checkIn(dialogId) {
  CacheService.getPrivateCache().put(dialogId + 'lastCheckIn', new Date().getTime());
}

function setDialogStatus(dialogId, status) {
  CacheService.getPrivateCache().put(dialogId, status);
}

function launchResetDialog(dialogId) {
  CacheService.getPrivateCache().remove(dialogId);
  return dialogId;
}

function getDialogStatus(dialogId) {
  var status = CacheService.getPrivateCache().get(dialogId);
  if (status != null) {
    return status;
  } else {
    var lastCheckIn = parseInt(CacheService.getPrivateCache().get(dialogId + 'lastCheckIn'));
    var now = new Date().getTime();
    if (now - lastCheckIn > DIALOG_TIMEOUT_SECONDS * 1000) {
      return 'lost';
    } else {
      return 'open';
    }
  }
}


function getSheetObj(ssId, sheetId) {  
  var ss = SpreadsheetApp.openById(ssId);
  var sheet = NVSL.getSheetById(sheetId, ss);
  var thisSheet = {};
  thisSheet.name = sheet.getName();
  thisSheet.id = sheetId;
  thisSheet.ssId = ssId;
  thisSheet.ssUrl = ss.getUrl();
  thisSheet.ssName = ss.getName();
  return thisSheet;
}


//spreadsheet functions
function getSheetObjs(ssId) {
  var ss = SpreadsheetApp.openById(ssId);
  var sheets = ss.getSheets();
  var sheetObjs = [];
  for (var i=0; i<sheets.length; i++) {
    var thisSheet = {};
    thisSheet.name = sheets[i].getName();
    thisSheet.id = sheets[i].getSheetId().toString();
    thisSheet.ssId = ssId;
    thisSheet.ssUrl = ss.getUrl();
    thisSheet.ssName = ss.getName();
    sheetObjs.push(thisSheet);
  }
  return sheetObjs;
}



function getSheetHeaders(sheetObj) {
  var ss = SpreadsheetApp.openById(sheetObj.ssId);
  var sheet = NVSL.getSheetById(sheetObj.id, ss);
  var headers = [];
  if (sheet.getLastColumn()>0) {
    var headersRange = sheet.getRange(1, 1, 1, sheet.getLastColumn())
    headers = NVSheetConverter.init(Session.getScriptTimeZone(), Session.getActiveUserLocale()).convertRange(headersRange)[0];
  }
  var scrubbedHeaders = [];
  for (var i=0; i<headers.length; i++) {
    if (headers[i]!=="") {
      scrubbedHeaders.push(headers[i]);
    }
  }
  return scrubbedHeaders;
}


function getDataPreview(sheetObj, selectedHeader) {
  var stringArray = [];
  try {
    if (!sheetObj) {
      return [];
    }
    if ((!sheetObj.ssId)||(!sheetObj.id)) {
      return [];
    }
    var ss = SpreadsheetApp.openById(sheetObj.ssId);
    var sheet = NVSL.getSheetById(sheetObj.id, ss);
    var headers = [];
    if (sheet.getLastColumn() > 0) {
      var headersRange = sheet.getRange(1, 1, 1, sheet.getLastColumn())
      headers = NVSheetConverter.init(Session.getScriptTimeZone(), Session.getActiveUserLocale()).convertRange(headersRange)[0];
    }
    var dataCol = headers.indexOf(selectedHeader)+1;
    var data = [["No data in column"]];
    if ((dataCol!=0)&&(sheet.getLastRow()>0)) {
      var dataRange = sheet.getRange(1, dataCol, sheet.getLastRow(), 1);
      data = NVSheetConverter.init(Session.getScriptTimeZone(), Session.getActiveUserLocale()).convertRange(dataRange);
    }
    for (var i=0; i< ((data.length>10) ? 10 : data.length); i++) {
      stringArray.push(data[i][0].toString())
    }
  } catch(err) {
    err = logErr_(err);
    throw err;
  }
    return stringArray;
}


function getDataByRangeId(rangeId, preview) {
  if (rangeId.indexOf("roster") === -1) {
    var rangeObj = getRangeById(rangeId);
    var stringArray = [];
    try {
      var ss = SpreadsheetApp.openById(rangeObj.ssId);
      var sheet = NVSL.getSheetById(rangeObj.sheetId, ss);
      var headers = [];
      if (sheet.getLastColumn() > 0) {
        var headersRange = sheet.getRange(1, 1, 1, sheet.getLastColumn())
        headers = NVSheetConverter.init(Session.getScriptTimeZone(), Session.getActiveUserLocale()).convertRange(headersRange)[0];
      }
      var dataCol = headers.indexOf(rangeObj.colHeading)+1;
      var data = [["No data in column"]];
      if ((dataCol!=0)&&(sheet.getLastRow()>0)) {
        var dataRange = sheet.getRange(2, dataCol, sheet.getLastRow()-1, 1);
        data = NVSheetConverter.convertRange(dataRange);
      }
      if (preview) {
        for (var i=0; i< ((data.length>10) ? 10 : data.length); i++) {
          if (data[i][0]!=='') {
            stringArray.push(data[i][0].toString())
          }
        }
      } else {
        for (var i=0; i<data.length; i++) {
          if (data[i][0]!=='') {
            stringArray.push(data[i][0].toString());
          }
        }
      }
    } catch(err) {
      stringArray = 'Unavailable';
    }
  } else {
    stringArray = buildRosterList(rangeId);
  }
  return stringArray;
}


function getRangeById(rangeId) {
  var ranges = getNamedRanges();
  for (var i=0; i<ranges.length; i++) {
    if (rangeId == ranges[i].rangeId) {
      return ranges[i];
    }
  }
  return;
}


function getNamedRanges() {
  var documentProperties = PropertiesService.getDocumentProperties();
  var namedRanges = documentProperties.getProperty('namedRanges');
  if (namedRanges) {
    namedRanges = JSON.parse(namedRanges);
  } else {
    namedRanges = [];
  }
  namedRanges.sort(function(a, b){
    if(a.name.toLowerCase() < b.name.toLowerCase()) return 1;
    if(a.name.toLowerCase() > b.name.toLowerCase()) return -1;
    return 0;
  })
  return namedRanges;
}


function getQuestionJobs(form, optFormId) {
  try {
    var documentProperties = PropertiesService.getDocumentProperties();
  } catch(err) {
    logErrInfo_(catchToString_(err));
    return [];
  }
  var questionJobs = documentProperties.getProperty('questionJobs');
  if (questionJobs) {
    questionJobs = JSON.parse(questionJobs);
  } else {
    questionJobs = [];
  }
  return questionJobs;
}


function deleteAllNamedRanges() {
  var documentProperties = PropertiesService.getDocumentProperties();
  var namedRanges = documentProperties.deleteProperty('namedRanges');
  return;
}




//Saves a new named range and returns its (very likely) unique id
//If rangeId is included, it will update the existing range
function saveNamedRange(ssId, sheetId, colHeading, rangeName, questionId, rangeId) {
  var documentProperties = PropertiesService.getDocumentProperties();
  var namedRanges = documentProperties.getProperty('namedRanges');
  var existingRangeIds = [];
  if (namedRanges) {
    namedRanges = JSON.parse(namedRanges);
    for (var i=0; i<namedRanges.length; i++) {
      existingRangeIds.push(parseInt(namedRanges[i].rangeId));
    }
  } else {
    namedRanges = [];
  }
  var thisRangeId = '';
  if (!rangeId) {
    var namedRange = {};
    namedRange.ssId = ssId;
    namedRange.sheetId = sheetId;
    namedRange.colHeading = colHeading;
    namedRange.name = rangeName;
    namedRange.rangeId = parseInt(Math.random() * 1e9).toString();
    while (existingRangeIds.indexOf(namedRange.rangeId) !== -1) {
      namedRange.rangeId = parseInt(Math.random() * 1e9).toString();
    }
    thisRangeId = namedRange.rangeId;
    namedRanges.push(namedRange);
  } else {
    var thisRangeIndex = existingRangeIds.indexOf(parseInt(rangeId));
    if (thisRangeIndex !== -1) {
      namedRanges[thisRangeIndex].ssId = ssId;
      namedRanges[thisRangeIndex].sheetId = sheetId;
      namedRanges[thisRangeIndex].colHeading = colHeading;
      namedRanges[thisRangeIndex].name = rangeName;
      namedRanges[thisRangeIndex].rangeId = rangeId;
      thisRangeId = rangeId;
    }
  }
  namedRanges = JSON.stringify(namedRanges);
  documentProperties.setProperty('namedRanges', namedRanges);
  saveQuestionJob(questionId.toString(), thisRangeId.toString());
  return thisRangeId;
}


//saves a question job given question id and range id
function saveQuestionJob(questionId, rangeId) {
  var documentProperties = PropertiesService.getDocumentProperties();
  var questionJobs = getQuestionJobs();
  var eligibleQuestionIds = getEligibleQuestionIds();
  var update = false;
  for (var i=0; i<questionJobs.length; i++) {
    if (questionJobs[i].questionId == questionId) { //if an existing question job was found for this question, update rangeId
      questionJobs[i].rangeId = rangeId;
      update = true;
    }
    if (eligibleQuestionIds.indexOf(questionId) === -1) { //remove question job if it is no longer assigned to an eligible type
      questionJobs.splice(i, 1);
    }
  }
  if (!update) {
    questionJobs.push({questionId: questionId.toString(), rangeId: rangeId.toString()});
  }
  questionJobs = JSON.stringify(questionJobs);
  documentProperties.setProperty('questionJobs', questionJobs);
  var status = populateFormQuestion(questionId, rangeId);
  status = questionId + "||" + status;
  return status;
}


function deleteQuestionJob(questionId) {
  var documentProperties = PropertiesService.getDocumentProperties();
  var questionJobs = documentProperties.getProperty('questionJobs');
  if (questionJobs) {
    questionJobs = JSON.parse(questionJobs);
  } else {
    questionJobs = [];
  }
  var qIdArray = [];
  for (var i=0; i<questionJobs.length; i++) {
    qIdArray.push(questionJobs[i].questionId.toString());
  }
  var thisQIndex = qIdArray.indexOf(questionId.toString());
  if (thisQIndex !== -1) {
    questionJobs.splice(thisQIndex, 1);
  }
  questionJobs = JSON.stringify(questionJobs);
  documentProperties.setProperty('questionJobs', questionJobs);
  var uiGlobal = getUiGlobal();
  clearFormQuestion(questionId);
  return uiGlobal;
}


function deleteQuestionJobsAndRange(rangeId) {
  var documentProperties = PropertiesService.getDocumentProperties();
  var questionJobs = documentProperties.getProperty('questionJobs');
  if (questionJobs) {
    questionJobs = JSON.parse(questionJobs);
  } else {
    questionJobs = [];
  }
  for (var i=0; i<questionJobs.length; i++) {
    if (questionJobs[i].rangeId == rangeId) {
      deleteQuestionJob(questionJobs[i].questionId)
    }
  }
  var namedRanges = getNamedRanges();
  var rIdArray = [];
  for (var i=0; i<namedRanges.length; i++) {
    rIdArray.push(namedRanges[i].rangeId);
  }
  var thisRangeIndex = rIdArray.indexOf(rangeId);
  if (thisRangeIndex !== -1) {
    namedRanges.splice(thisRangeIndex, 1);
  }
  namedRanges = JSON.stringify(namedRanges);
  documentProperties.setProperty('namedRanges', namedRanges);
  var uiGlobal = getUiGlobal();
  return uiGlobal;
}


function getEligibleQuestionIds() {
  var formQuestions = getActiveFormQuestions();
  var questionIds = [];
  for (var i=0; i<formQuestions.length; i++) {
    if (APPROVED_TYPES.indexOf(formQuestions[i].type) !== -1) {
      questionIds.push(formQuestions[i].id.toString());
    }
  }
  return questionIds;
}

/**
* Gets the OAuth token authorized for the scopes that this add-on requests.
* Since we want to access documents in Google Drive, we include a call to
* a DriveApp method to generate this access scope.
*
* @return The authorized OAuth token
*/
function getOAuthToken() {
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();;
}

function closeUi() {
  return;
}



