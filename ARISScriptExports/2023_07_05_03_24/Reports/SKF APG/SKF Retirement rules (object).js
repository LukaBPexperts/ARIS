/* DEBUG VARIABLES */

var DEBUG = false;
var SETPROPERTY = false;
var DEBUG_OBJECTS = "O:d08cc970-a525-11e7-3310-005056a95aef;";
/* DEBUG SETTINGS */

var g_log = [];

function main() {
    try {
        var retirementResult = [];
        var incomingRetirementResult = null;
        var incomingRetirementResultString = String(Context.getProperty("retirement_result"));
        if (incomingRetirementResultString!=null) incomingRetirementResult = incomingRetirementResultString.split(";");
        
        var objDefList = ArisData.getSelectedObjDefs();
        if (objDefList.length==0) {
            if (SETPROPERTY) Context.setProperty("OBJECT_LIST",DEBUG_OBJECTS);
            var objectListString = Context.getProperty("OBJECT_LIST");  
            if (objectListString==null) addErrorLog("Property 'OBJECT_LIST' is empty!");
            objDefList = getObjectsFromAPGString(objectListString);
        }
        var modelList = [];
        for each (var objDef in objDefList) { 
            var lockedModels = getLockedModels(objDef);
            
            var retired = false;
            if ( isAllowedObjDef(objDef)) {
                if (lockedModels.length>0) {
                    addErrorLog("Object '" + objDef.Name(-1)  + "', retirement not executed, object is occurring on these locked models:\n" + itemListToString(lockedModels, "\n", Constants.AT_NAME, true));   
                } else {
                    retireObject(objDef);    
                    var objOccList = objDef.OccList();
                    objOccList = filterOccListByModelType(objOccList, MT_INTERFACE_FLOW_DGM);
                    for each (var objOcc in objOccList) {
                        modelList.push(objOcc.Model());    
                    }
                    deactivateObjOccList(objOccList, AT_EAM_CLASSIFICATION,  ATV_RETIRED);
                    retired = true;
                }
            }
            retirementResult.push(objDef.GUID()+":" + String(retired));                             
        }  
        var g_modelList_RuleA = [];
        var g_modelList_RuleB = [];
        var g_modelList_RuleC = [];
        var g_modelList_RuleD = [];
        var g_modelList_RuleE = [];
        var g_modelList_RuleF = [];
        
        for each (var model in modelList) {   
                
            var interfaceObjOcc = getCentralInterface(model) ;
            if (interfaceObjOcc!=null) {
                var senderOccList = getConnectedObjOccsByObjDefType(interfaceObjOcc, null, [CT_SEND_DATA], null, Constants.EDGES_IN) ;
                senderOccList = filterRelevantObjOccs(senderOccList);
                
                var receiverOccList = getConnectedObjOccsByObjDefType(interfaceObjOcc, null, [CT_SEND_DATA], null, Constants.EDGES_OUT) ;
                receiverOccList = filterRelevantObjOccs(receiverOccList); 
                 
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
                if ((senderOccList.length>1 && checkAttributeValueInItemList(senderOccList, AT_EAM_CLASSIFICATION,  ATV_RETIRED)) 
                        || (receiverOccList.length>1 && checkAttributeValueInItemList(receiverOccList, AT_EAM_CLASSIFICATION,  ATV_RETIRED) )) {
                    retire = true;
                    g_modelList_RuleC.push(model);               
                } else {
                    deactivate = true;
                }
                
                // RULE D - If the interface flow diagram contains an Application Service that has EAM Classification attribute value = “Retired”
                var applicationServices = model.ObjOccListFilter(OT_APPLICATION_SERVICE);
                if (geItemsByAttributeValue(applicationServices, AT_EAM_CLASSIFICATION,  ATV_RETIRED).length>0) {
                    retire = true;
                    g_modelList_RuleD.push(model);
                }
                
                // RULE E - the interface flow diagram contains a Logical Data Object that has EAM Classification attribute value = “Retired”
                var logicalDataOccList = model.ObjOccListFilter(OT_LOGICAL_DATA);
                if (geItemsByAttributeValue(logicalDataOccList, AT_EAM_CLASSIFICATION,  ATV_RETIRED).length>0) {
                    retire = true;
                    g_modelList_RuleE.push(model);
                }                
                
                // RULE F - If the interface flow diagram contains an Interface object that has EAM Classification attribute value = “Retired”, 
                var interfaceOccList = model.ObjOccListFilter(OT_INTERFACE);
                if (geItemsByAttributeValue(interfaceOccList, AT_EAM_CLASSIFICATION,  ATV_RETIRED).length>0) {
                    retire = true;
                    g_modelList_RuleF.push(model);
                }                         
                
                
                if (retire==true || deactivate==true) {
                    if (model.canWrite(true)==false) {
                        addErrorLog("Model is read only, cannot be retired: '" + getItemTitle(model) + "'");
                        continue;
                    }
                    if (deactivate==true) {
                        var appOccList = model.ObjOccListFilter ( OT_APPLICATION_SERVICE);
                        var logicalDataObjectList = model.ObjOccListFilter ( OT_LOGICAL_DATA);
                        deactivateObjOccList(senderOccList.concat(receiverOccList.concat(appOccList.concat(logicalDataObjectList))), AT_EAM_CLASSIFICATION,  ATV_RETIRED);
                    }               
                    if (retire==true) {
                        retireModel(model, interfaceObjOcc); 
                    } 
                }
                
            }                            
        }  
        if (modelList.length>0) {
            g_log.push("The following models were retired as follows:");
            g_log.push("Retirement rule A executed for: \n" + (g_modelList_RuleA.length>0?itemListToString(g_modelList_RuleA, "\n", Constants.AT_NAME, true): "  no model"));
            g_log.push("Retirement rule B executed for: \n" + (g_modelList_RuleB.length>0?itemListToString(g_modelList_RuleB, "\n", Constants.AT_NAME, true): "  no model"));
            g_log.push("Retirement rule C executed for: \n" + (g_modelList_RuleC.length>0?itemListToString(g_modelList_RuleC, "\n", Constants.AT_NAME, true): "  no model"));
            g_log.push("Retirement rule D executed for: \n" + (g_modelList_RuleD.length>0?itemListToString(g_modelList_RuleD, "\n", Constants.AT_NAME, true): "  no model"));
            g_log.push("Retirement rule E executed for: \n" + (g_modelList_RuleE.length>0?itemListToString(g_modelList_RuleE, "\n", Constants.AT_NAME, true): "  no model"));
            g_log.push("Retirement rule F executed for: \n" + (g_modelList_RuleF.length>0?itemListToString(g_modelList_RuleF, "\n", Constants.AT_NAME, true): "  no model"));
        } else {
            g_log.push("No model was retired");
        }
         
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
    Context.setProperty("retirement_result" , retirementResult.join(";"))
    if (DEBUG||SETPROPERTY) showPropertiesInDialog();
}

function checkRetirementResult(resultList, objDefGUID) {
    if (resultList) {
        for each (var resultPair in resultList) {
            [guid, result] = resultPair.split(":");
            if (guid.equals(objDefGUID)) {
                return result.equals("true");
            }
        }
    } else {
        return true;
    }
    return null;
}

function getLockedModels(objDef) {
    var objOccList = objDef.OccList();
    var modelList = [];
    for each (var objOcc in objOccList) {
        if (!objOcc.Model().canWrite(true)) {
            modelList.push(objOcc.Model());
        }
    }
    modelList = ArisData.Unique(modelList);
    return modelList;
}

function filterRelevantObjOccs(senderOccList) {
    var returnArray = [];
    for each (var objOcc in senderOccList) {
        var objDef = objOcc.ObjDef();
        if ( objDef.TypeNum()==OT_APPLICATION_SERVICE ) {
            if (isApplicationService(objDef)) {
                returnArray.push(objOcc);
            }
        } else if (objDef.TypeNum()==OT_ROLE ) {
            if (isTradingPartner(objDef)) {
                returnArray.push(objOcc);
            }
        } else {
            returnArray.push(objOcc);    
        }      
    }
    return returnArray;
}

function isAllowedObjDef(objDef) {
    var filter = ArisData.ActiveFilter();
    if (objDef.TypeNum()==OT_APPLICATION_SERVICE) {
        if (objDef.Attribute(AT_APP_SERVICE_TYPE,-1).getValue().equals(ATV_APP_SERVICE_TYPE_MS) ) {
            return true;
        } else {
            addErrorLog("Object: '" + objDef.Name(-1) + "' - context is of type '" + filter.ObjTypeName(OT_APPLICATION_SERVICE) + "', but the attribute '" + filter.AttrTypeName ( AT_APP_SERVICE_TYPE ) + "' has not allowed value: '" + objDef.Attribute(AT_APP_SERVICE_TYPE,-1).getValue() + "', object will be ignored." ); 
            return false;
        }
    } else if (objDef.TypeNum()==OT_ROLE) {
        if (objDef.Attribute(AT_TECHNICAL_NAME_TRADING_PARTNER,-1).IsMaintained() ) {
            return true; 
        } else {
            addErrorLog("Object: '" + objDef.Name(-1) + "' - context is of type '" + filter.ObjTypeName(OT_ROLE) + "', but attribute '" + filter.AttrTypeName ( AT_TECHNICAL_NAME_TRADING_PARTNER ) + "' is not maintained, object will be ignored." ); 
            return false;            
        }
    } else if (containsInt([OT_APPLICATION_INSTANCE, OT_INTERFACE, OT_LOGICAL_DATA], objDef.TypeNum() ))  {
        return true;
    } else {
        return false;
    }
}

function getObjectsFromAPGString(string) {
    var returnArray = [];
    if (string==null) return [];
    
    var splitCommaList = string.split(";");
    for each (var splitComma in splitCommaList) {
        var splitTypeAndGUID = splitComma.split(":");
        if (splitTypeAndGUID.length==2) {
            var foundItem = null;
            if (splitTypeAndGUID[0].equals("O")) {
                foundItem = ArisData.getActiveDatabase().FindGUID(splitTypeAndGUID[1], Constants.CID_OBJDEF);
            }   
            if (foundItem!=null && foundItem.IsValid()) {
                returnArray.push(foundItem);  
            } else {
                addErrorLog("Item could not be found for " + splitTypeAndGUID);
            }
        }
    }
    return returnArray;
}
function getInterfacesFromAPGString(string) {
    var returnArray = [];
    if (string==null) return [];
    
    var splitCommaList = string.split(";");
    for each (var splitComma in splitCommaList) {
        var splitTypeAndGUID = splitComma.split(":");
        if (splitTypeAndGUID.length==2) {
            var foundItem = null;
            if (splitTypeAndGUID[0].equals("O")) {
                foundItem = ArisData.getActiveDatabase().FindGUID(splitTypeAndGUID[1]);
            }   
            if (foundItem!=null && foundItem.IsValid()) {
                if (containsInt([OT_INTERFACE],foundItem.TypeNum()) ) {
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



var g_errorLog = [];

function deactivateObjOccList(objOccList, attTypeNum, checkedValue) {
    for each (var item in objOccList) {
        if (item.ObjDef().Attribute(attTypeNum,-1).getValue().equals(checkedValue) ) {
            var changed = item.setActive(false);
            if (changed!=null && changed==false) addErrorLog("Could not change active flag on object '" + getItemTitle( item.ObjDef() )+ "', model: '" + getItemTitle( item.Model() ) + "'");  
        }
    }
}

function checkAttributeValueInItemList(itemList, attTypeNum, checkedValue) {
    var returnValue = true;
    for each (var item in itemList) {
        if (!item.ObjDef().Attribute(attTypeNum,-1).getValue().equals(checkedValue) ) {
           returnValue = false;    
        }
    }
    return returnValue;
}
function geItemsByAttributeValue(itemList, attTypeNum, checkedValue) {
    var returnValue = [];
    for each (var item in itemList) {
        if (item.ObjDef().Attribute(attTypeNum,-1).getValue().equals(checkedValue) ) {
           returnValue.push(item);    
        }
    }
    return returnValue;
}


function showPropertiesInDialog() {
    var returnStringList = [];
    for each (var key in Context.getProperties()) {
        returnStringList.push(key + ":\n" + Context.getProperty(key));
    }
    
	var result = Dialogs.showDialog(new errorDialog(returnStringList), Constants.DIALOG_TYPE_WIZARD, "Log");
}

function retireModel(model, interfaceObjOcc) {
    if (model.Attribute(AT_EAM_GOV_STAGE_MODEL, -1).getValue().equals(ATV_RETIRED)) {
        //addErrorLog("Model retirement skipped, model is already retired: '" + getItemTitle(model) + "'");
    } else {
        var changed = model.Attribute(AT_EAM_GOV_STAGE_MODEL, -1).setValue(ATV_RETIRED); // Change EAM Governance Stage attribute value for the model to “Retired”
        
        var changed = setNamePrefix(model, PREFIX_RETIRED);//Change the model name to have “Retired – “ as prefix.
        if (changed!=null && changed==false) addErrorLog("Could not set prefix for model " + getItemTitle(model) );
    
        moveToRetiredGroup(model); 
        
        interfaceObjOcc.setActive(false);
        
        //Move the interface object to applicable sub-folder in “Y Retired” folder
        retireObject(interfaceObjOcc.ObjDef());
    }
    
}

function moveToRetireGroupOld(item) {
    var db = ArisData.getActiveDatabase();    
    var retiredPathBase = RETIRED_PATH.replace("[root]",db.RootGroup().Path(-1));

    //Move the item to applicable sub-folder in “Y Retired” folder
    var itemSubPath = item.Group().Path(-1).substring(db.RootGroup().Path(-1).length(), item.Group().Path(-1).length()) ;
    var retiredItemGroup = getCreateGroup(retiredPathBase + itemSubPath);
    var changed = item.ChangeGroup ( retiredItemGroup);
    if (changed!=null && changed==false) addErrorLog("Could not change path to '" + retiredPathBase + itemSubPath + "' of model '" + getItemTitle(item)  + "'"  );      
}

function moveToRetiredGroup(item) {
    var db = ArisData.getActiveDatabase();    
    var retiredPathBase = RETIRED_PATH.replace("[root]",db.RootGroup().Path(-1));
    
    if (left(item.Group().Path(-1), retiredPathBase.length).equals(retiredPathBase)) {
        addErrorLog("Relocation of item to archive folder skipped, item is already in archived folder: '" + getItemTitle(item) + "'");
    } else {
        var itemSubPath = "";
        if (item.KindNum()==Constants.CID_MODEL) {
            itemSubPath = "Models";
            itemSubPath = itemSubPath + "\\"  +  item.Type() + "s"
            if (item.TypeNum()==MT_INTERFACE_FLOW_DGM) {
                var folderNameA2A = item.Type() + "s" + " " + INTERFACE_A2A_STRING;
                var folderNameB2B = item.Type() + "s" + " " + INTERFACE_B2B_STRING;
                if (isB2B_A2AModel(item, folderNameA2A)) {
                    itemSubPath = itemSubPath + "\\" + folderNameA2A;
                } else if (isB2B_A2AModel(item, folderNameB2B)) {
                    itemSubPath = itemSubPath + "\\" + folderNameB2B;
                }
            }
        } else {
            itemSubPath = itemSubPath + item.Type() + "s";
            if (item.TypeNum()==OT_INTERFACE) {
                var folderNameA2A = item.Type() + "s" + " " + INTERFACE_A2A_STRING;
                var folderNameB2B = item.Type() + "s" + " " + INTERFACE_B2B_STRING;
                
                if (isB2B_A2AModel(item, folderNameA2A)) {
                    itemSubPath = itemSubPath + "\\" + folderNameA2A;
                } else if (isB2B_A2AModel(item, folderNameB2B)) {
                    itemSubPath = itemSubPath + "\\" + folderNameB2B;
                }    
            }               
        }
    
    
        
        var targetPath = retiredPathBase + "\\" + itemSubPath;
            
        var retiredItemGroup = getCreateGroup(targetPath);
        var changed = item.ChangeGroup ( retiredItemGroup);
        if (changed!=null && changed==false) addErrorLog("Could not change path to '" + targetPath + "' of model '" + getItemTitle(item)  + "'"  );   
    }
}

function left(str, n){
    var str = String(str);
    if (n <= 0)
        return "";
    else if (n > str.length)
        return str;
    else
        return str.substring(0,n);
}

function isA2A_B2BObject(item, foderName) {  
    return item.Group().Path(-1).indexOf(folderName)>-1;    
}

function isB2B_A2AModel(model, folderName) {
    return model.Group().Path(-1).indexOf(folderName)>-1;
}

function retireObject(interfaceObjDef) {
    /*if (interfaceObjDef.Attribute(AT_EAM_CLASSIFICATION, -1).getValue().equals(ATV_RETIRED)) {
        addErrorLog("Object retirement skipped, object is already retired: '" + getItemTitle(interfaceObjDef) + "'");
    } else {*/
        var changed = interfaceObjDef.Attribute(AT_EAM_CLASSIFICATION,-1).setValue(ATV_RETIRED); //Change the EAM Classification attribute value for the Interface object to “Retired”
        if (changed!=null && changed==false) addErrorLog("EAM Classification attribute could not be set to 'Retired', object " + getItemTitle(interfaceObjDef) );
        
        var changed = setNamePrefix(interfaceObjDef, PREFIX_RETIRED);//Change the name on the interface object to have “Retired – “ as prefix”.
        if (changed!=null && changed==false) addErrorLog("Could not set prefix for object " + getItemTitle(interfaceObjDef));   
        
        //Move the interface object to applicable sub-folder in “Y Retired” folder
        moveToRetiredGroup(interfaceObjDef);
    //}
}

function setNamePrefix(model, prefix) {
    var name = model.Name(-1);
    if (name.indexOf(prefix)!=0 ) {
        return model.Attribute(Constants.AT_NAME,-1).setValue(prefix + name);                
    }
    return null;
}


function getCentralInterface(model) {
    var objOccList = model.ObjOccListBySymbol (ST_INTERFACE);        
    if (objOccList.length>0) {
        return objOccList[0];    
    }
    return null;
}

function addErrorLog(msg){
    g_errorLog.push(msg);
}

function getConnectedObjOccsByObjDefType(p_objOcc, p_objTypesList, p_includedCxnTypes, p_excludedCxnTypes, p_direction) {
    var returnArray = new Array();
    if (p_direction!=null) {
        var cxnList = p_objOcc.Cxns(p_direction);          
    } else {
        var cxnList = p_objOcc.CxnOccList();      
    }
    for (var i=0; i<cxnList.length; i++) {
        var cxnOcc = cxnList[i];
        var cxnDef = cxnOcc.CxnDef();  
        var targetObjOcc = null;
        if (cxnDef.SourceObjDef().IsEqual(p_objOcc.ObjDef()) ) {
            targetObjOcc = cxnOcc.TargetObjOcc();
        } else {
            targetObjOcc = cxnOcc.SourceObjOcc();
        }
        var isIncluded = true;         
        if (p_includedCxnTypes==null) {
            isIncluded = true;    
        } else {
            if (containsInt(p_includedCxnTypes, cxnDef.TypeNum())) {
                isIncluded = true;
            } else {
                isIncluded = false;
            }
        }
        var isExcluded = false;          
        if (p_excludedCxnTypes==null) {
            isExcluded = false;    
        } else {
            if (containsInt(p_excludedCxnTypes, cxnDef.TypeNum()) )  {
                isExcluded = true;
            } else {
                isExcluded = false;
            }
        }        
        var isAllowedObjType = false;
        if (p_objTypesList==null) {
            isAllowedObjType = true;
        } else {
            isAllowedObjType = containsInt(p_objTypesList, targetObjOcc.ObjDef().TypeNum());
        }
        if (isIncluded && !isExcluded && isAllowedObjType) {
            returnArray.push(targetObjOcc);
        }
    }
    return returnArray;
}

function itemListToString(p_itemlist, p_separator, p_attType, p_withGUID) {
    var attType = (p_attType==null)? Constants.AT_NAME : p_attType;
    var loc =  Context.getSelectedLanguage();      
    
    var returnString = "";
    for (var i=0; i<p_itemlist.length; i++) {
        var item = p_itemlist[i];
            
        if (i!=0) {
            returnString += p_separator;
        }
        
        if (item.KindNum()==Constants.CID_OBJOCC) {
            returnString += "    - " + item.ObjDef().Attribute(attType, loc, true).GetValue(true); 
        } else {
            if ((typeof p_attType === 'string') && p_attType.equals("GUID")) {
                returnString += item.GUID();    
            } else {
                returnString += "    - " + item.Attribute(attType, loc, true).GetValue(true);                
            }
            if (p_withGUID) returnString += " (" + item.GUID() + ")";
        }

    }
    
    //APG
    //if (returnString.length==0) returnString = p_separator;    
    return returnString;
} 
function getException(ex) {
    return ex.message + " (line: "+ ex.lineNumber + ", file: " + ex.fileName + ") ";
}

function getItemTitle(item) {
    return item.Name(-1) + " (" + item.GUID() + ")";
}
/** 
* getCreateGroup() finds or creates a group if the requested path doesn't exist 
* 
* @param sPath        path to check 
*/   

function getCreateGroup(sPath, nLoc) { 
    if (nLoc==null) nLoc = -1;
    var iPos=-1; 
    var sParentPath; 
    var sGroupName; 

    var idsParentGroup=null; 
    var idsGroup=null; 

    idsGroup=ArisData.getActiveDatabase().Group(sPath, nLoc)       

    if (!idsGroup.IsValid()){ 
        iPos=sPath.lastIndexOf('\\'); 
        if (iPos>=0){ 
            sParentPath=sPath.substring(0, iPos); 
            sGroupName=sPath.substring(iPos+1); 
            idsParentGroup= getCreateGroup(sParentPath, nLoc); 
        }else{ 
            sGroupName=sPath; 
            idsParentGroup=ArisData.getActiveDatabase().RootGroup(); 
        } 
        if (idsParentGroup.IsValid()){ 
            idsGroup=idsParentGroup.CreateChildGroup(sGroupName, nLoc); 
        } 

    }else{ 
        return idsGroup; 
    } 
    return idsGroup;         
} 

function errorDialog(p_msgList) {
	this.msgList = p_msgList;
	this.getPages = function() {
		var iDialogTemplate1 = Dialogs.createNewDialogTemplate(1100, 300, "Errors have occurred");
		iDialogTemplate1.TextBox(10, 10, 1080, 330, "TXT_EDIT_1",1);
		return [iDialogTemplate1];
	}

	this.init = function(aPages) {
		//use this function also to store the page data locally (for example to access it in "onClose")
		aPages[0].setFocusedElement("TXT_EDIT_1");
		aPages[0].getDialogElement("TXT_EDIT_1").setText(p_msgList.join("\n\n"));
		aPages[0].getDialogElement("TXT_EDIT_1").setEnabled ( true );        
        
	}

	this.isInValidState = function(pageNumber)	{
		return true;
	}
}

function filterOccListByModelType(objOccList, modelTypeNum) {
    var returnArray = [];
    for each (var objOcc in objOccList) {
        if (objOcc.Model().TypeNum()==modelTypeNum) {
            returnArray.push(objOcc);
        }
    }
    return returnArray;
}


main();