/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

var modelToValidate = ArisData.getSelectedModels()[0];

main();
function main() {
    try {
        if ((modelToValidate == null) && (!getBoolPropertyValue("SILENT_MODE"))) {
            Dialogs.MsgBox(getString("MODEL_DELETED"), Constants.MSGBOX_BTN_OK, Context.getScriptInfo(Constants.SCRIPT_NAME));
            return;
        }  
        
        var validationResults = validateModel(modelToValidate);
        
        if ((validationResults.length == 0) && (!getBoolPropertyValue("SILENT_MODE"))){
            Dialogs.MsgBox(getString("CHECKING_SUCCESS"), Constants.MSGBOX_BTN_OK, Context.getScriptInfo(Constants.SCRIPT_NAME));
            return;
        }
        
        setValidationResultToProperties(validationResults);
        showValidationResults(modelToValidate, validationResults);
    } catch (error) {
        if (!getBoolPropertyValue("SILENT_MODE")){
            Dialogs.MsgBox(getString("UNEXPECTED_ERROR"), Constants.MSGBOX_BTN_OK, Context.getScriptInfo(Constants.SCRIPT_NAME));
        }
    }
}

/**
* Validates the model and returns the results
*/
function validateModel(modelToValidate) {
    var modelInterchangeComponent = Context.getComponent("ModelInterchange");
    return modelInterchangeComponent.validateBpmnDiagramForExport(modelToValidate);
}

/**
* Shows the validation results in the model
* Works only if the model is opened and focused while script executing
*/
function showValidationResults(model, validationResults) {
    model.openModel();
    var objOccs = model.ObjOccList();
    var cxnOccs = model.CxnOccList();
    
    var modelErrors = "";
    for (var j = 0; j < validationResults.length; j++) {
        var currentResult = validationResults[j];
        
        for (var i = 0; i < objOccs.length; i++) {
            if (currentResult.getOID().compareTo(objOccs[i].ObjectID()) == 0) {
                addResultMessageToObject(objOccs[i], currentResult);
                continue;
            }
        }
        
        for (var i = 0; i < cxnOccs.length; i++) {
            if (currentResult.getOID().compareTo(cxnOccs[i].ObjectID()) == 0) {
                addResultMessageToObject(cxnOccs[i], currentResult);
                continue;
            }
        }
        
        if (currentResult.getOID().compareTo(model.ObjectID()) == 0) {
            var entryMessages = currentResult.getMessages();
            modelErrors = concatStringsForInformationDisplay(entryMessages);
        }
        
        if ((modelErrors != "") && (!getBoolPropertyValue("SILENT_MODE"))){
            Dialogs.MsgBox(modelErrors, Constants.MSGBOX_BTN_OK, Context.getScriptInfo(Constants.SCRIPT_NAME));
        }
    }
}

/**
* Adds the message from the result to the object so its displayed after the validation process
*/
function addResultMessageToObject(object, result) {
    var errorMessages = result.getMessages();
    var outgoingMessage = concatStringsForInformationDisplay(errorMessages);            
    modelToValidate.setTemporaryUserInfo(object, result.getMessageLevel(), outgoingMessage);
}

/**
* Concats the strings so they are shown in seperate lines
*/
function concatStringsForInformationDisplay(strings) {
    var outgoingMessage = "<html>";
    
    for (var i = 0; i < strings.length; i++) {
        outgoingMessage += "<br>- " + strings[i];
    }
    
    outgoingMessage += "</html>";
    
    return outgoingMessage;
}


/**
* Sets the validation results to the properties
* This is needed for testing this script and may be removed if not needed
*/
function setValidationResultToProperties(result) {
    var errorList = new java.util.ArrayList();
    
    for (var i = 0; i < result.length; i++) {
        var entry = new java.util.ArrayList();
        entry.add(result[i].getOID());
        entry.add(result[i].getMessageLevel());
        entry.add(result[i].getMessages().length);
        
        errorList.add(entry);
    }
    
    Context.setProperty("ERROR_LIST", errorList);
}


/**
* Returns the named property as boolean
*/
function getBoolPropertyValue(p_sPropKey) {
    var property = Context.getProperty(p_sPropKey);
    
    if (property != null) {
        return (StrComp(property, "true") == 0);
    }
    
    return false;
}