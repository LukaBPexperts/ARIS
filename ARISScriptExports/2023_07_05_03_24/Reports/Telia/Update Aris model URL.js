//----------Editable variables--------------
var g_publishingDBName = "_TC_Publishing"; //To check that report is not run in wrong database
var g_workingDBName = "_TeliaCompany_WORKING";
var g_adminUser = "arisservice"; //user account that performs all chenges
var g_adminPwd = "arisservice"; //password for user account
var g_filterGuid = "dd838074-ac29-11d4-85b8-00005a4053ff"; //Entire method
//----------------------------------------------

var modelLinkPublishingAttNum = ArisData.ActiveFilter().UserDefinedAttributeTypeNum("cf3816a0-06b9-11ee-2420-00505682bf5e") //Model Link - Published version
var modelLinkWorkingAttNum = ArisData.ActiveFilter().UserDefinedAttributeTypeNum("c2624d50-06ba-11ee-2420-00505682bf5e") //Model Link - Working version		

var nLocale = Context.getSelectedLanguage()

//Log into current DB as administrator
var publishingDB = ArisData.openDatabase(g_publishingDBName, g_adminUser, g_adminPwd, g_filterGuid, nLocale);
publishingDB.setAutoTouch(false);
var workingDB = ArisData.openDatabase(g_workingDBName, g_adminUser, g_adminPwd, g_filterGuid, nLocale);
workingDB.setAutoTouch(false);

var allModels = publishingDB.Find(Constants.SEARCH_MODEL)

for (var i = 0; i < allModels.length; i++) {
    var model = allModels[i]
    //update link att on model in publishing DB
    var linkPublishedDB = getPortalLink(model, publishingDB)
    var updatePublishingAtt = model.Attribute(modelLinkPublishingAttNum, nLocale).setValue(linkPublishedDB)

    //check if the model exist in Working DB, if yes update link
    var modelWorkingDB = workingDB.FindGUID(model.GUID())
    if (modelWorkingDB.IsValid()) {
        var linkWorkingDB = getPortalLink(modelWorkingDB, workingDB)
        //update link att on model in working DB
        var updatePublishingAtt = modelWorkingDB.Attribute(modelLinkPublishingAttNum, nLocale).setValue(linkPublishedDB)
        var updateWorkingAtt = modelWorkingDB.Attribute(modelLinkWorkingAttNum, nLocale).setValue(linkWorkingDB)
        //update working DB link on model in publishing DB
        var updateWorkingAtt = model.Attribute(modelLinkWorkingAttNum, nLocale).setValue(linkWorkingDB)
    }
}

publishingDB.setAutoTouch(true);
publishingDB.close()
workingDB.setAutoTouch(true);
workingDB.close()

function getPortalLink(model, oDatabase) {
    //http://<servername>/#<tenant>/item/c.<factSheet>.<DBName>.<GUID>.<Version>

    //var oDatabase = ArisData.getActiveDatabase();
    var sTenantName = ArisData.getTenantName();
    var serverName = oDatabase.ServerName();
    var sDBName = oDatabase.Name(nLocale);
    sDBName = sDBName.replace(" ", "%20");
    //var mGUID = model.GUID()
    var mGUID = compressGuid(model.GUID())

    //var linkURL = "http://" + serverName + "/#" + sTenantName + "/item/c." + sDBName + "." + mGUID + ".-1"
    var linkURL = "https://" + serverName + "/#" + sTenantName + "/item/c." + sDBName + "." + mGUID
    return linkURL
}

function compressGuid(guid) {
    var code = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_";
    var compressedGuid = "";
    var bucket = 0;
    var bits = 0;
    for (var i = 0; i < guid.length(); i++) {
        
        var rValue = guid.charCodeAt(i);
        if (rValue >= 48 && rValue <=57)
            fValue = rValue - 48;
        else if (rValue >= 97 && rValue <= 102)
            fValue = 10 + rValue - 97;
        else if (rValue >= 65 && rValue <= 90)
            fValue = 10 + rValue - 65;
        else            
            continue;
 
        bucket = (bucket << 4) + fValue;
        bits+=4;
        if (bits >= 6) {
            bits-=6;
            sValue = bucket >> bits;
            bucket = bucket ^ (sValue << bits);
            compressedGuid += code.charAt(sValue);
        }
    }
    if (bits > 0) {
        compressedGuid += code.charAt(bucket << (6 - bits));
    }
    return compressedGuid;
}