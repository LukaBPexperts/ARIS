const CT_SEND_DATA = Constants.CT_SENDS_3;

const ST_INTERFACE = sym("1154cc71-4355-11e6-6fc6-005056a95aef");

const OT_APPLICATION_SERVICE = Constants.OT_SCRN;
const OT_APPLICATION_INSTANCE =	Constants.OT_APPL_SYS_TYPE;
const OT_LOGICAL_DATA =	Constants.OT_TECH_TRM;
const OT_INTERFACE = Constants.OT_CLS;
const OT_ROLE = Constants.OT_PERS_TYPE;
const MT_INTERFACE_FLOW_DGM = mt("5cac0681-1073-11e6-6fc6-005056a95aef");

const ATV_RETIRED = "Retired";
const PREFIX_RETIRED = "RETIRED – ";

const AT_EAM_GOV_STAGE_MODEL = att("8f8039b1-6d98-11e5-15b5-b4b52f5aab7a");
const AT_EAM_CLASSIFICATION = att("f8dc7e01-8dcd-11e5-6fc6-005056a95aef");

const RETIRED_PATH = "[root]\\Y Retired";

const INTERFACE_A2A_STRING = "A2A";
const INTERFACE_B2B_STRING = "B2B";

const AT_APP_SERVICE_TYPE = att("be4d3d11-d549-11e5-6fc6-005056a95aef"); 	
const ATV_APP_SERVICE_TYPE_MS = "Message Mapping";

const AT_TECHNICAL_NAME_TRADING_PARTNER = att("e90d49f1-bf96-11e5-6fc6-005056a95aef");


function isApplicationService(objDef) {
    return (objDef.TypeNum()==OT_APPLICATION_SERVICE && objDef.Attribute(AT_APP_SERVICE_TYPE,-1).getValue().equals(ATV_APP_SERVICE_TYPE_MS) );    
}
function isTradingPartner(objDef) {
    return (objDef.TypeNum()==OT_ROLE && objDef.Attribute(AT_TECHNICAL_NAME_TRADING_PARTNER,-1).IsMaintained() ); 
}

function containsString(p_stringList, p_string ) {
    for (var i=0; i<p_stringList.length; i++) {
        if (p_stringList[i].equals(p_string)) return true;
    }
    return false;
}

function containsInt(p_intList, p_int ) {
    for (var i=0; i<p_intList.length; i++) {
        if (p_intList[i]==p_int) return true;
    }
    return false;
}


/* function att()
*  Transfer the specified attribute type GUID to a current attribute type number. 
*  returns: attribute type number
*/    
function att(typeNumOrGUID) {
    var returnTypeNum = null;
    try {
        var returnTypeNum = ArisData.getActiveDatabase().ActiveFilter().UserDefinedAttributeTypeNum(typeNumOrGUID);           
    } catch (e) {};
    try {
        if ((returnTypeNum == null) && (typeNumOrGUID.length!=36)) {
            returnTypeNum = parseInt(typeNumOrGUID);
        }
    } catch (e) {}
    return returnTypeNum;
}
/* function sym()
*  Transfer the specified attribute type GUID to a current attribute type number. 
*  returns: attribute type number
*/    
function sym(typeNumOrGUID) {
    var returnTypeNum = null;
    try {
        var returnTypeNum = ArisData.getActiveDatabase().ActiveFilter().UserDefinedSymbolTypeNum (typeNumOrGUID);           
    } catch (e) {};
    try {
        if ((returnTypeNum == null) && (typeNumOrGUID.length!=36)) {
            returnTypeNum = parseInt(typeNumOrGUID);
        }
    } catch (e) {}
    return returnTypeNum;
}

function atv(attGUID, attValueGUID) {
    return ArisData.getActiveDatabase().ActiveFilter().UserDefinedAttributeValueTypeNum ( attGUID, attValueGUID );
}

/* function mt()
*  Transfer the specified attribute type GUID to a current model type number. 
*  returns: attribute type number
*/    
function mt(typeNumOrGUID) {
    var returnTypeNum = null;
    try {
        var returnTypeNum = ArisData.getActiveDatabase().ActiveFilter().UserDefinedModelTypeNum(typeNumOrGUID);           
    } catch (e) {};
    try {
        if ((returnTypeNum == null) && (typeNumOrGUID.length!=36)) {
            returnTypeNum = parseInt(typeNumOrGUID);
        }
    } catch (e) {}
    return returnTypeNum;
}

