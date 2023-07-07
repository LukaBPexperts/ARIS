/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

// BLUE-14292   Cover all database (optionally)
// AGG-14252    Adapt tenant
 
function DLG_OPTIONS(p_server, p_tenant, p_bAllDBs) {
    this.server = p_server;
    this.tenant = p_tenant;
    this.bAllDBs = p_bAllDBs;
} 

var options = getOptions();
if (options != null && (options.server != "" || options.tenant != "")) {

    updateADSLinks(ArisData.getActiveDatabase(), options);      // Update active database
    
    if (options.bAllDBs) {        
        var dbNames = ArisData.GetDatabaseNames();
        for (var i in dbNames) {
            var dbName = dbNames[i];
            if (isActiveDB(dbName)) continue;

            // Uses the same user account name, language and filter as the initial database login from which the report was started. 
            var db = ArisData.openDatabase(dbName);
            if (db != null) {

                updateADSLinks(db, options);                    // Update further databases
                db.close();
            } else {
                Context.writeLog(formatstring1(getString("ERROR_LOG2"), dbName));        
            }
        }
    }
}

/******************************************************************************************************/ 

function updateADSLinks(oDB, options) {

    oDB.setAutoTouch(false);      // No touch !!! 
    
    var langList = oDB.LanguageList();  
    
    var allObjects = oDB.Find(Constants.SEARCH_OBJDEF); 
    for (var i in allObjects) 
    {
        var oObject = allObjects[i];
        updateAttr(oObject, Constants.AT_ADS_LINK_1);
        updateAttr(oObject, Constants.AT_ADS_LINK_2);
        updateAttr(oObject, Constants.AT_ADS_LINK_3);
        updateAttr(oObject, Constants.AT_ADS_LINK_4);
    }
   
    var allModels = oDB.Find(Constants.SEARCH_MODEL);    
    for (var i in allModels) 
    {
        var oObject = allModels[i];
        updateAttr(oObject, Constants.AT_ADS_LINK_1);
        updateAttr(oObject, Constants.AT_ADS_LINK_2);
        updateAttr(oObject, Constants.AT_ADS_LINK_3);
        updateAttr(oObject, Constants.AT_ADS_LINK_4);
    }
    
    function updateAttr(oObject, atn) {
        for (var i in langList) {
            var locale = langList[i].LocaleId();
            
            var linkAttr = oObject.Attribute(atn, locale);
            var linkStr = ""+linkAttr.getValue();

            var bDoUpdate = false;
            
            // update server
            if (options.server != "") {
                var index = linkStr.indexOf("/ads");
                if (index < 0) {
                    index = linkStr.indexOf("/documents");
                }
                if (index >= 0) {
                    linkStr = options.server + linkStr.substring(index);
                    bDoUpdate = true;
                }
            }
            
            // update tenant
            if (options.tenant != "") {
                var sPattern = "tenantid=";                
                var indexStart = linkStr.indexOf(sPattern);
                if (indexStart >= 0) {
                    indexStart = indexStart + sPattern.length
                    var linkSubStr = linkStr.substring(indexStart);
                    
                    var indexEnd = linkSubStr.indexOf("&");
                    if (indexEnd >= 0) {
                        
                        linkStr = linkStr.substring(0, indexStart) + options.tenant + linkSubStr.substring(indexEnd);
                        bDoUpdate = true;
                    }
                }
            }

            if (options.tenant != "") {
                var sPattern = "tenant=";                
                var indexStart = linkStr.indexOf(sPattern);
                if (indexStart >= 0) {
                    indexStart = indexStart + sPattern.length
                    var linkSubStr = linkStr.substring(indexStart);
                    
                    var indexEnd = linkSubStr.indexOf("&");
                    if (indexEnd >= 0) {
                        
                        linkStr = linkStr.substring(0, indexStart) + options.tenant + linkSubStr.substring(indexEnd);
                        bDoUpdate = true;
                    }
                }
            }
                            
            if (bDoUpdate) {
                var bUpdated = linkAttr.setValue(linkStr);
				
                if (!bUpdated) Context.writeLog(formatstring3(getString("ERROR_LOG"), oObject.GUID(), atn, locale));
            }
        }

    }
}

function isActiveDB(dbName) {
    return StrComp(ArisData.getActiveDatabase().Name(-1), dbName) == 0;
}

function getOptions() {
    return Dialogs.showDialog(new optionsDialog(), Constants.DIALOG_TYPE_ACTION, "Select report options");

    function optionsDialog() {
        var bOkClicked = false;
        
        var sDlgText = getString("DLG_TEXT");
        if (sDlgText == "") sDlgText = "ARIS Server URL (e.g. https://newhost:8180):"

        this.getPages = function() {
            var iDialogTemplate = Dialogs.createNewDialogTemplate(500, 100, "");
            
            iDialogTemplate.GroupBox(10, 15, 480, 60, getString("DLG_SERVER"));
            iDialogTemplate.Text(25, 25, 450, 15, sDlgText);          
            iDialogTemplate.TextBox(25, 40, 450, 21, "TXT_SERVER");
            
            iDialogTemplate.GroupBox(10, 90, 480, 60, getString("DLG_TENANT"));
            iDialogTemplate.Text(25, 100, 450, 15, getString("DLG_TEXT3"));          
            iDialogTemplate.TextBox(25, 115, 450, 21, "TXT_TENANT");
            
            iDialogTemplate.CheckBox(5, 165, 350, 21, getString("DLG_TEXT2"), "CHECK_DB");
            
            return [iDialogTemplate];
        }
        
        // optional
        this.init = function(aPages){
        }
        // mandatory
        this.isInValidState = function(pageNumber) {
            return true;
        }
        // optional
        this.onClose = function(pageNumber, bOk) {
            bOkClicked = bOk;
        }
        // optional
        this.getResult = function() {
            if (bOkClicked) {
                var page = this.dialog.getPage(0);
                var server = page.getDialogElement("TXT_SERVER").getText();
                var tenant = page.getDialogElement("TXT_TENANT").getText();                
                var bAllDBs = page.getDialogElement("CHECK_DB").isChecked();
                
                return new DLG_OPTIONS(server, tenant, bAllDBs);
            }
            return null;
        }
    }
}
        
