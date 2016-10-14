function logRepopulatedFormQuestions_() {
  NVAddOns.log("Repopulated%20Form%20Questions", scriptName, scriptTrackingId);
}


function logRepeatInstall_() {
  NVAddOns.log("Repeat%20Install", scriptName, scriptTrackingId);
}

function logFirstInstall_() {
  NVAddOns.log("First%20Install", scriptName, scriptTrackingId);
}

// Call this function from within the first major UI Setup step.
// Multiple calls to this function will not result in multiple install analytics

function setSid_() { 
  var docProperties = PropertiesService.getDocumentProperties();
  var userProperties = PropertiesService.getUserProperties();
  var scriptNameLower = scriptName.toLowerCase();
  var sid = docProperties.getProperty(scriptNameLower + "_sid");
  if (sid == null || sid == "")
  {
    incrementNumUses_();
    var dt = new Date();
    var ms = dt.getTime();
    var ms_str = ms.toString();
    docProperties.setProperty(scriptNameLower + "_sid", ms_str);
    var uid = userProperties.getProperty(scriptNameLower + "_uid");
    if (uid) {
      logRepeatInstall_();
    } else {
      logFirstInstall_();
      userProperties.setProperty(scriptNameLower + "_uid", ms_str);
    }      
  }
}


function incrementNumUses_() {
  try {
    var numFormRangerUses = PropertiesService.getUserProperties().getProperty('numFormRangerUses');
    if (parseInt(numFormRangerUses)) {
      numFormRangerUses = parseInt(numFormRangerUses) + 1;
    } else {
      numFormRangerUses = 1;
    }
    PropertiesService.getUserProperties().setProperty('numFormRangerUses', numFormRangerUses);
  } catch(err) {
    var errInfo = catchToString_(err);
    logErrInfo_(errInfo);
    SpreadsheetApp.getUi().alert("Oops! An error has occurred and logged so the New Visions CloudLab team can look into it...");
  }
}

