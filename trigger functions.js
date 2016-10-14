function unsetTriggerServerSide() {
  var formId = PropertiesService.getDocumentProperties().getProperty('formId');
  var form = FormApp.openById(formId);
  var triggers = ScriptApp.getUserTriggers(form);
     var found = false;
    for (var i=0; i<triggers.length; i++) {
      if ((triggers[i].getHandlerFunction() === "populateAllFormQuestions")&&(triggers[i].getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT)) {
        ScriptApp.deleteTrigger(triggers[i]);
        PropertiesService.getDocumentProperties().deleteProperty('triggerState');
        return "removed trigger";
      }
    }
    PropertiesService.getDocumentProperties().deleteProperty('triggerState');
    return "no trigger found";
}


function unsetTimeTriggerServerSide() {
  var formId = PropertiesService.getDocumentProperties().getProperty('formId');
  var form = FormApp.openById(formId);
  var triggers = ScriptApp.getUserTriggers(form);
     var found = false;
    for (var i=0; i<triggers.length; i++) {
      if ((triggers[i].getHandlerFunction() === "populateAllFormQuestions")&&(triggers[i].getEventType() === ScriptApp.EventType.CLOCK)) {
        ScriptApp.deleteTrigger(triggers[i]);
        PropertiesService.getDocumentProperties().deleteProperty('timeTriggerState');
        return "removed trigger";
      }
    }
    PropertiesService.getDocumentProperties().deleteProperty('timeTriggerState');
    return "no trigger found";
}


function setTriggerServerSide() {
  try {
    var triggerState = triggerIsSet();
    var user = Session.getActiveUser().getEmail();
    if (!triggerState) {
      triggerState = {user: user};
      var formId = PropertiesService.getDocumentProperties().getProperty('formId');
      var form = FormApp.openById(formId);
      var triggers = ScriptApp.getUserTriggers(form);
      var found = false;
      for (var i=0; i<triggers.length; i++) {
        if ((triggers[i].getHandlerFunction() === "populateAllFormQuestions")&&(triggers[i].getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT)) {
          found = true;
        }
      }
      if (!found) {
        ScriptApp.newTrigger('populateAllFormQuestions').forForm(form).onFormSubmit().create();
        PropertiesService.getDocumentProperties().setProperty('triggerState', JSON.stringify(triggerState));
        return "success";
      } else {
        PropertiesService.getDocumentProperties().setProperty('triggerState', JSON.stringify(triggerState));
        return "already set";
      }
    } else if (triggerState.user === user) {
      return "set by this user";
    } else {
      return triggerState.user;
    }
  } catch(err) {
    err = logErr_(err);
    throw err;
  }
}


function setTimeTriggerServerSide() {
  try {
    var triggerState = timeTriggerIsSet();
    var user = Session.getActiveUser().getEmail();
    if (!triggerState) {
      triggerState = {user: user};
      var formId = PropertiesService.getDocumentProperties().getProperty('formId');
      var form = FormApp.openById(formId);
      var triggers = ScriptApp.getUserTriggers(form);
      var found = false;
      for (var i=0; i<triggers.length; i++) {
        if ((triggers[i].getHandlerFunction() === "populateAllFormQuestions")&&(triggers[i].getEventType() === ScriptApp.EventType.CLOCK)) {
          found = true;
        }
      }
      if (!found) {
        ScriptApp.newTrigger('populateAllFormQuestions').timeBased().everyHours(1).create();
        PropertiesService.getDocumentProperties().setProperty('timeTriggerState', JSON.stringify(triggerState));
        return "success";
      } else {
        PropertiesService.getDocumentProperties().setProperty('timeTriggerState', JSON.stringify(triggerState));
        return "already set";
      }
    } else if (triggerState.user === user) {
      return "set by this user";
    } else {
      return triggerState.user;
    }
  } catch(err) {
    err = logErr_(err);
    throw err;
  }
}



function getTriggerState() {
  try {
    var docProperties = PropertiesService.getDocumentProperties();
    var formId = docProperties.getProperty('formId');
    if (!formId) {
      var form = FormApp.getActiveForm();
      formId = form.getId();
      docProperties.setProperty('formId', formId);
    }
    var triggerState = triggerIsSet();
    var user = Session.getActiveUser().getEmail();
    if (!triggerState) {
      return triggerState;
    } else if (triggerState.user === user) {
      return "this_user";
    } else {
      return triggerState.user;
    }
  } catch(err) {
    err = logErr_(err);
    throw err;
  }
}



function getTimeTriggerState() {
  try {
    var docProperties = PropertiesService.getDocumentProperties();
    var formId = docProperties.getProperty('formId');
    if (!formId) {
      var form = FormApp.getActiveForm();
      formId = form.getId();
      docProperties.setProperty('formId', formId);
    }
    var triggerState = timeTriggerIsSet();
    var user = Session.getActiveUser().getEmail();
    if (!triggerState) {
      return triggerState;
    } else if (triggerState.user === user) {
      return "this_user";
    } else {
      return triggerState.user;
    }
  } catch(err) {
    err = logErr_(err);
    throw err;
  }
}


function triggerIsSet() {
  var triggerState = PropertiesService.getDocumentProperties().getProperty('triggerState');
  if (triggerState) {
    triggerState = JSON.parse(triggerState);
    return triggerState;
  } else {
    return false;
  }
}


function timeTriggerIsSet() {
  var triggerState = PropertiesService.getDocumentProperties().getProperty('timeTriggerState');
  if (triggerState) {
    triggerState = JSON.parse(triggerState);
    return triggerState;
  } else {
    return false;
  }
}


function deleteAppUserTriggers(form) {
  var triggers = ScriptApp.getUserTriggers(form);
  for (var i=0; i<triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}


function checkAuthStatus() {
  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  // Check if the actions of the trigger require authorizations
  // that have not been supplied yet -- if so, warn the active
  // user via email (if possible).  This check is required when
  // using triggers with add-ons to maintain functional triggers.
  if (authInfo.getAuthorizationStatus() == ScriptApp.AuthorizationStatus.REQUIRED) {
    var props = PropertiesService.getDocumentProperties();
    // Re-authorization is required. In this case, the user
    // needs to be alerted that they need to reauthorize; the
    // normal trigger action is not conducted, since it authorization
    // needs to be provided first. Send at most one
    // 'Authorization Required' email a day, to avoid spamming
    // users of the add-on.
    var lastAuthEmailDate = props.getProperty('lastAuthEmailDate');
    var today = new Date().toDateString();
    if (lastAuthEmailDate != today) {
      if (MailApp.getRemainingDailyQuota() > 0) {
        var html = HtmlService.createTemplateFromFile('AuthorizationEmail');
        html.url = authInfo.getAuthorizationUrl();
        html.addonTitle = addonTitle;
        var message = html.evaluate();
        MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
                          'Authorization Required',
                          message.getContent(), {
                            name: addonTitle,
                            htmlBody: message.getContent()
                          }
                         );
        logAuthEmailSent_();
      }
      props.setProperty('lastAuthEmailDate', today);
    }
    return false;
  } else {
    return true;
  }
}

