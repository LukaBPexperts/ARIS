/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

var sModelid = Context.getProperty("modelId"); 
var sMetadataKeys = Context.getProperty("keys"); 
var aMetadataKeys = new Array();
if (sMetadataKeys != null) 
{
    aMetadataKeys = sMetadataKeys.split("\b");
}

var sMetadataValues = Context.getProperty("values"); 
var aMetadataValues = new Array();
if (sMetadataValues != null) 
{
    aMetadataValues = sMetadataValues.split("\b");
}

var mapKeyValue = new java.util.HashMap();
for (var t = 0; t < aMetadataKeys.length; t++)
{
    mapKeyValue.put(new java.lang.String(aMetadataKeys[t]), new java.lang.String(aMetadataValues[t]));
}

var comp = Context.getComponent("webMethodsIntegration");
var model = ArisData.getActiveDatabase().FindGUID(sModelid);
if (model.IsValid())
{
    var result = comp.setARISMetaData(model, mapKeyValue);
                      
    if (!result.succeeded())
    {
        handleResult(result.getErrorMsg());
    }
    else
    {
        var out = new XMLFormattedOut();
        out.setSuccess(true); 
        out.write();
    }
}
else
{
    handleResult("Model with GUID " + sModelid + " not found in " + ArisData.getActiveDatabase().toString());
}

function handleResult(p_errorMessage) {
	var errOut = new XMLFormattedOut();
	errOut.setSuccess(false, p_errorMessage);
	errOut.write();
}