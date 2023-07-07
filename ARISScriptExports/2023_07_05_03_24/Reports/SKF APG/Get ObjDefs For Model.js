
// FOR Testing
//Context.setProperty("out_object_guids","O:cec96570-a65e-11e5-38c7-080027650386;O:cec96572-a65e-11e5-38c7-080027650386;O:cec96575-a65e-11e5-38c7-080027650386");
//Context.setProperty("out_sup_object_guid","O:68c39660-a81b-11e5-38c7-080027650386");
  



//----------------------------------------------------------------------------------------------------------------
// ARIS Script - Get ObjDefs For Model
//
// (c) Ingo Opetnik
// Software GmbH Austria
//
// Version 1.0.0 (01.12.2015|Ingo Opetnik)
//	- Basic functionality
//
// Version 1.1.0 (26.01.2016|Ingo Opetnik)
//	- added "in_attribute_status"
//  - added "in_attribute_status_value"
//
//----------------------------------------------------------------------------------------------------------------  
           
    var g_nLoc = Context.getSelectedLanguage();
    var g_hmTypeMapping = new java.util.HashMap();
    var g_sInModelGUID = Context.getProperty("in_model_guid");
    var g_sObjTypes = Context.getProperty("in_object_types");
//    var g_sInModelGUID = "M:500fcec1-ba88-11e5-38c7-080027650386";
//    var g_sObjTypes = "OT_APPL_SYS_TYPE;OT_ORG_UNIT";    
    var g_nStatusAttrTypeNum = HLP_getAttributeTypeNum(Context.getProperty("in_attribute_status"));
    var g_sStatusAttrValue = Context.getProperty("in_attribute_status_value");
    var g_sOutObjDefsGUID = "";
    var g_sOutSupObjDefGUID = "";
    
    main();

    
  

//------------------------------------------------------------------------------------------------------------------------
//------------------------------------------------------------------------------------------------------------------------
//------------------------------------------------------------------------------------------------------------------------
//------------------------------------------------------------------------------------------------------------------------
//------------------------------------------------------------------------------------------------------------------------

//*
//* 'Main' routine of the script
//*

function main() {   
    
    getObjDefsForModel();
        
    Context.setProperty("out_object_guids", g_sOutObjDefsGUID);
    Context.setProperty("out_sup_object_guid", g_sOutSupObjDefGUID);

}

function getObjDefsForModel() {           

    // remove M: char
    if (g_sInModelGUID.indexOf("M:") == 0) {
        g_sInModelGUID = g_sInModelGUID.substring(2);
    }
    
    var oModel = ArisData.getActiveDatabase().FindGUID(g_sInModelGUID);
    if (oModel == null || !oModel.IsValid()) {
        g_sOutSupObjDefGUID = "Could not find model!";
        return;    
    }
    
    // determine objdefs
    var anObjTypeNums = getObjTypeNums();
    var aoObjDefs = oModel.ObjDefListByTypes(anObjTypeNums); 
    aoObjDefs = ArisData.sort(aoObjDefs, Constants.AT_NAME, g_nLoc);
    var bFound = false;
    for (var i = 0; i < aoObjDefs.length; i++) {
        
        var oAttr = aoObjDefs[i].Attribute(g_nStatusAttrTypeNum, g_nLoc);
        if (oAttr.IsMaintained() && oAttr.getValue().equals(g_sStatusAttrValue)) {
            continue;
        }
        
        bFound = true;        
        if (String(g_sOutObjDefsGUID).length > 0) {
            g_sOutObjDefsGUID += ";";
        }
        g_sOutObjDefsGUID += "O:" + aoObjDefs[i].GUID();               
    }
    
    if (!bFound) {
        g_sOutObjDefsGUID = ";";    
    }
    
    // determine superior objdefs
    var aoSupObjDefs = oModel.SuperiorObjDefs();
    if (aoSupObjDefs.length == 1) {
        g_sOutSupObjDefGUID = "O:" + aoSupObjDefs[0].GUID();
    }
    else if (aoSupObjDefs.length == 0) {
    //no superior object
        g_sOutSupObjDefGUID = ";";        
    }
    else {
// Multiple superior objects
        g_sOutSupObjDefGUID = ";";        
    }
}                

function getObjTypeNums() {
    var anRetObjTypeNums = new Array();
    var asTemp = g_sObjTypes.split(";");
    
    for (var i = 0; i < asTemp.length; i++) {
        anRetObjTypeNums.push(eval("Constants." + asTemp[i]));
    }
    
    return anRetObjTypeNums;
}

function HLP_getAttributeTypeNum(sValue) {
    var returnValue = null;

    try {
        if (String(sValue).length==36) {
            var attTypeNumFromGUID = null
            try {
                attTypeNumFromGUID = ArisData.getActiveDatabase().ActiveFilter().UserDefinedAttributeTypeNum(sValue);
            } catch (e) {}
            //returns null if the GUID was not found (!)
            returnValue = attTypeNumFromGUID;
        } else {
            returnValue = sValue;
        }
    } catch (e) {
        returnValue = sValue;
    }

    return returnValue;
}