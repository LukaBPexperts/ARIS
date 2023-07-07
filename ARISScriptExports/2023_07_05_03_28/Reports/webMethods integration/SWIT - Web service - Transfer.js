/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

var name = Context.getProperty("processName");
var guid = Context.getProperty("parentId");
var xpdl = Context.getProperty("process.xmlProcess");
var metadata  = Context.getProperty("process.metadata");
var group = ArisData.getActiveDatabase().FindGUID(guid);

var out = new XMLFormattedOut();
if(group == null)
    out.setSuccess(false, "Error: parent group not found.");
else{
    var comp = Context.getComponent("webMethodsIntegration");
    if(comp == null)
        out.setSuccess(false, "Error: webMethods integration component not found.");
    else{
        try{
            var guid = comp.createProcess(group, name,xpdl, metadata);
            out.addElement("id", guid);
	        out.setSuccess(true);
        }catch(error){
            out.setSuccess(false, error.toString());
        }
    }
}
out.write();
