function populateAllFormQuestions(e) {
  try {
    if (checkAuthStatus()) {
      var questionsUnableToPopulate = [];
      try {
        var form = e ? e.source : FormApp.getActiveForm();  //assume it's running on form trigger or from ui
        var formId = form.getId();
      } catch(err) {
      }
      if (!form) {  //assume it's running on time trigger     
        try {
          formId = PropertiesService.getDocumentProperties().getProperty('formId');
          form = FormApp.openById(formId);
        } catch(err) {
          logErrInfo_(catchToString_(err));
          return;
        }
      }
      var questionJobs = getQuestionJobs(form, formId);
      for (var i=0; i<questionJobs.length; i++) {
        var thisRangeId = questionJobs[i].rangeId;
        var populated = populateFormQuestion(questionJobs[i].questionId, thisRangeId, form);
        if (populated === "Unable to populate") {
          questionsUnableToPopulate.push(questionJobs[i]);
        }
      }
      try {
        logRepopulatedFormQuestions_();
      } catch(err) {
      }
    }
    return questionsUnableToPopulate;
  } catch(err) {
    err = logErr_(err);
    throw err;
  }
}



function populateFormQuestion(questionId, rangeId, form) {
  var theseOptions = getDataByRangeId(rangeId);
  if (theseOptions!== "Unavailable") {
    if (!form) {
      var form = FormApp.getActiveForm();
      if (!form) {
        var formId = PropertiesService.getDocumentProperties().getProperty('formId');
        form = FormApp.openById(formId);
      }
    }
    try {
      var thisQuestion = form.getItemById(questionId);
      if (!thisQuestion) {      //garbage collection on deleted questions
        deleteQuestionJob(questionId);
        return;
      }
      var type = thisQuestion.getType();
      if (type == "LIST") {
        thisQuestion.asListItem().setChoiceValues(theseOptions);
      }
      if (type == "MULTIPLE_CHOICE") {
        var isBranching = false;
        var existingChoices = thisQuestion.asMultipleChoiceItem().getChoices();
        var existingChoiceValues = [];
        for (var i=0; i<existingChoices.length; i++) {
          var choiceValue = existingChoices[i].getValue();
          var navType = existingChoices[i].getPageNavigationType() ? existingChoices[i].getPageNavigationType().toString() : '';
          if ((navType)&&(navType !== "CONTINUE")) {
            isBranching = true;
          }
          existingChoiceValues.push(choiceValue);
        }
        if (isBranching) {
          var choices = [];
          for (var i=0; i<theseOptions.length; i++) {
            var thisChoiceIndex = existingChoiceValues.indexOf(theseOptions[i]);
            if (thisChoiceIndex === -1) {
              var thisChoice = thisQuestion.asMultipleChoiceItem().createChoice(theseOptions[i], FormApp.PageNavigationType.CONTINUE);
              choices.push(thisChoice)
            } else {
              choices.push(existingChoices[thisChoiceIndex]);
            }
          }
          thisQuestion.asMultipleChoiceItem().setChoices(choices);
        } else {
          thisQuestion.asMultipleChoiceItem().setChoiceValues(theseOptions);
        }
      }
      if (type == "CHECKBOX") {
        thisQuestion.asCheckboxItem().setChoiceValues(theseOptions);
      }
      if (type == "GRID") {
        var existingRows = thisQuestion.asGridItem().getRows();
        var hasNewOptions = false;
        for (var i=0; i<theseOptions.length; i++) {
          if (existingRows.indexOf(theseOptions[i]) === -1) { 
            hasNewOptions = true;
          }
        }
        if (hasNewOptions) {
          thisQuestion.asGridItem().setRows(theseOptions);
        }
      }
    } catch(err) {
      Logger.log(err.message);
    }
    return;
  } else {
    return "Unable to populate";
  }
}




function clearFormQuestion(questionId) {
  var questionJobs = getQuestionJobs();
  var form = FormApp.getActiveForm();
   if (!form) {
      if (!form) {
        var formId = PropertiesService.getDocumentProperties().getProperty('formId');
        form = FormApp.openById(formId);
      }
    }
  try {
      var thisQuestion = form.getItemById(questionId);
      var type = thisQuestion.getType();
      if (type == "LIST") {
        thisQuestion.asListItem().setChoiceValues(['']);
      }
      if (type == "MULTIPLE_CHOICE") {
        thisQuestion.asMultipleChoiceItem().setChoiceValues(['']);
      }
      if (type == "CHECKBOX") {
        thisQuestion.asCheckboxItem().setChoiceValues(['']);
      }
    } catch(err) {
      Logger.log(err.message);
    }
}
