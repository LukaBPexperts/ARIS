var DEBUG = false;

var g_log = [];

function main() {
    try {
        var modelList = ArisData.getSelectedModels();
        if (modelList.length==0) {
            if (DEBUG) Context.setProperty("MODEL_LIST","M:b87477c1-1195-11e8-3310-005056a95aef");
            var modelListString = Context.getProperty("MODEL_LIST");  
            if (modelListString==null) addErrorLog("Property 'MODEL_LIST' is empty!");
            modelList = getFlowModelsFromAPGString(modelListString);
        }

        var g_modelList_RuleA = [];
        var g_modelList_RuleB = [];
        var g_modelList_RuleC = [];
        var g_modelList_RuleD = [];
        for each (var model in modelList) {   
                
            var interfaceObjOcc = getCentralInterface(model) ;
            if (interfaceObjOcc!=null) {
                var senderOccList = getConnectedObjOccsByObjDefType(interfaceObjOcc, null, [CT_SEND_DATA], null, Constants.EDGES_IN) ;
                var receiverOccList = getConnectedObjOccsByObjDefType(interfaceObjOcc, null, [CT_SEND_DATA], null, Constants.EDGES_OUT) ;
                
                var retire = false;
                var deactivate = false;
                // RULE A
                if (receiverOccList.length==1 && receiverOccList[0].ObjDef().Attribute(AT_EAM_CLASSIFICATION,-1).getValue().equals(ATV_RETIRED)  ) {
                    retire = true;   
                    g_modelList_RuleA.push(model);
                } 
    
                // RULE B
                if (senderOccList.length==1 && senderOccList[0].ObjDef().Attribute(AT_EAM_CLASSIFICATION,-1).getValue().equals(ATV_RETIRED)  ) {
                    retire = true;
                    g_modelList_RuleB.push(model);
                }       
                
                // RULE C
                if ((senderOccList.length>1 || receiverOccList.length>1) && checkAttributeValueInItemList(senderOccList.concat(receiverOccList), AT_EAM_CLASSIFICATION,  ATV_RETIRED) ) {
                    retire = true;
                    g_modelList_RuleC.push(model);               
                } else {
                    deactivate = true;
                }
                
                // RULE D
                var applicationServices = model.ObjOccListFilter(Constants.OT_SCRN);
                if (checkAttributeValueInItemList(applicationServices, AT_EAM_CLASSIFICATION,  ATV_RETIRED)) {
                    retire = true;
                    g_modelList_RuleD.push(model);
                }
                
                if (retire==true || deactivate==true) {
                    if (model.canWrite(true)==false) {
                        addErrorLog("Model is read only, cannot be retired: '" + getItemTitle(model) + "'");
                        continue;
                    }
                    if (deactivate==true) {
                        var appOccList = model.ObjOccListFilter ( Constants.OT_SCRN);
                        var logicalDataObjectList = model.ObjOccListFilter ( Constants.OT_TECH_TRM);
                        deactivateObjOccList(senderOccList.concat(receiverOccList.concat(appOccList.concat(logicalDataObjectList))), AT_EAM_CLASSIFICATION,  ATV_RETIRED);
                    }               
                    if (retire==true) {
                        retireModel(model, interfaceObjOcc); 
                    } 
                }
                
            }                            
        }  
        
        g_log.push("Retirement rule A executed for: \n" + (g_modelList_RuleA.length>0?itemListToString(g_modelList_RuleA, "\n", Constants.AT_NAME, true): "  no model"));
        g_log.push("Retirement rule B executed for: \n" + (g_modelList_RuleB.length>0?itemListToString(g_modelList_RuleB, "\n", Constants.AT_NAME, true): "  no model"));
        g_log.push("Retirement rule C executed for: \n" + (g_modelList_RuleC.length>0?itemListToString(g_modelList_RuleC, "\n", Constants.AT_NAME, true): "  no model"));
        g_log.push("Retirement rule D executed for: \n" + (g_modelList_RuleD.length>0?itemListToString(g_modelList_RuleD, "\n", Constants.AT_NAME, true): "  no model"));
         
    } catch (ex) {
        addErrorLog(getException(ex));
    }
    
    
    if (g_errorLog.length>0) {
        Context.setProperty("error", true);    
        Context.setProperty("errorLog", g_errorLog.join("\n")); 
    } else {
        Context.setProperty("error", false);       
    }
    
    Context.setProperty("log", g_log.join("\n"));
    if (DEBUG) showPropertiesInDialog();
}


function getFlowModelsFromAPGString(string) {
    var returnArray = [];
    if (string==null) return [];
    
    var splitCommaList = string.split(";");
    for each (var splitComma in splitCommaList) {
        var splitTypeAndGUID = splitComma.split(":");
        if (splitTypeAndGUID.length==2) {
            var foundItem = null;
            if (splitTypeAndGUID[0].equals("M")) {
                foundItem = ArisData.getActiveDatabase().FindGUID(splitTypeAndGUID[1]);
            }   
            if (foundItem!=null && foundItem.IsValid()) {
                if (containsInt([MT_INTERFACE_FLOW_DGM],foundItem.TypeNum()) ) {
                    returnArray.push(foundItem);  
                } else {
                    addErrorLog( "Item for '" + splitTypeAndGUID + "' has different type: '" + foundItem.Type());
                }
            } else {
                addErrorLog("Item could not be found for " + splitTypeAndGUID);
            }
        }
    }
    return returnArray;
}

main();