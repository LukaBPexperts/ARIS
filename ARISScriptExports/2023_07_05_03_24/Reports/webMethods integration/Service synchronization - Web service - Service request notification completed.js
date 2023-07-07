/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

//informs AIS about a finished request

Context.setProperty("excel-formulas-allowed", false); //default value is provided by server setting (tenant-specific): "abs.report.excel-formulas-allowed"

var dbname = Context.getProperty("dbname");
var language = Context.getProperty("language");
var objectGuid = Context.getProperty("objectGuid");

// The requst type is "implemented" or "canceled". Currently there is no automatic handling in ARIS concerning 
// the type of the request.
var requestType = Context.getProperty("requestType");

var database = ArisData.getActiveDatabase();

// uddi: has to be remoced from the guid
if(objectGuid.startsWith("uddi:")){
    objectGuid = objectGuid.substring(5,objectGuid.length());
}

var sstObjDef = database.FindGUID(objectGuid);

var errOut = new XMLFormattedOut();
if (sstObjDef.IsValid()){
    var webMethodsComponent = Context.getComponent("webMethodsIntegration");
    var finishedSuccessful;
    var message;

    var result = webMethodsComponent.serviceRequstFinished(database, sstObjDef, language);  

    
    if(!result.succeeded()){
	    errOut.setSuccess(false, result.getErrorMsg());
	    
    }else{
        errOut.setSuccess(true, "Succesful")
    }
} else{
    errOut.setSuccess(false, getString("TEXT1") + "\n" + getString("TEXT2") + " " +database.Name(-1,true) + "\n" + getString("TEXT3") + " " + objectGuid );

}

errOut.write();
