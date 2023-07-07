/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

main();

function main()
{
    var modelGuid = Context.getProperty("modelId");
    var documentationZIP = Context.getProperty("zippedDocumentationProcess");
    
    if(modelGuid == null)
    {
        handleError("Model GUID is null");
        return;
    }

    var bpmnModel = ArisData.getActiveDatabase().FindGUID(modelGuid);
    
    if(bpmnModel == null || !bpmnModel.IsValid())
    {
        handleError("Model not found. The model does either not exist anymore or the user does not have the necessary access privileges."); 
        return;
    }
    
    
    try{
        var database = ArisData.getActiveDatabase();
        var webMethodsIntegrationComponent = Context.getComponent("webMethodsIntegration"); 
        webMethodsIntegrationComponent.storeDocumentaion(database, modelGuid, documentationZIP);
        var out = new XMLFormattedOut();
        out.setSuccess(true);
        out.write();
    } catch (e){
        handleError(e.getMessage());
    }
 
}

function handleError(p_errorMessage)
{
	var errOut = new XMLFormattedOut();
	errOut.setSuccess(false, p_errorMessage);
	errOut.write();
}