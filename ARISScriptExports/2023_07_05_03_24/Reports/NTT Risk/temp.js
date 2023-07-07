var g_oDatabase = ArisData.getActiveDatabase();
//var currentLng = Context.getSelectedLanguage();
var currentLng = -1
var oUser = ArisData.getActiveUser();

var startEnv =Context.getEnvironment()


/* CONSTANTS */
var c_requestTypeExistingRisk = 0;
var c_requestTypeNewRisk = 1;

var g_strUpdateRiskAPGGUID="c0a59c71-435a-11ed-30e7-001dd8d80275";

var c_riskRegisterObjectGUID = "e18acc20-0b03-11ec-51c7-d05099dfe77e" //starting Risk object GUID
var c_riskLogDatabaseName = "NTT Risk Version Logs"

var listPossibleStatuses = getPossibleValues(AT_NTT_RM_RISK_STATUS);
var listPossibleExecutiveManager = getPossibleValues(AT_NTT_RM_EXEC_MGMT_OWNERSHIP);
var listPossiblePrincipalRisk = getPossibleValues(AT_NTT_RM_PRINCIPAL_RISK_IMPACTED);
var listPossibleSubRisk = getPossibleValues(AT_NTT_RM_SUB_RISK_CATEGORY);
var listPossibleRiskTrend = getPossibleValues(AT_NTT_RM_RISK_TREND);
var listPossibleGDCEMEACountry = getPossibleValues(AT_NTT_RM_GDC_EMEA_AFFECTED_COUNTRY);
var listPossibleGDCEMEASite = getPossibleValues(AT_NTT_RM_GDC_EMEA_AFFECTED_SITE);
var listPossibleRiskReach = getPossibleValues(AT_NTT_RM_RISK_REACH);
var listPossibleRiskImpact = getPossibleValues(AT_NTT_RM_RISK_IMPACT);
var listPossibleRiskProbability = getPossibleValues(AT_NTT_RM_RISK_PROBABILITY);
var listPossibleRiskStrengthOfControls = getPossibleValues(AT_NTT_RM_STRENGTH_OF_CONTROLS);
var listPossibleRiskTreatmentStrategy = getPossibleValues(AT_NTT_RM_RISK_TREATMENT_STRATEGY);
var listPossibleRiskCanBeResolved = getPossibleValues(AT_NTT_RM_CAN_THIS_BE_FIXED);
var listPossibleRiskMitigationStatus = getPossibleValues(AT_NTT_RM_MITIGATION_STATUS);
var listPossibleDepartments = getPossibleValues(AT_NTT_RM_DEPARTMENT);

var oRiskRegisterLogic = new RiskRegisterLogic("NTT\\Risk Register Configuration", "Risk Register-strength of controls.xlsx", "Risk Register-strength of controls", "Backend");

//SHOW DIALOG to select update Type: NEW or EXISTING RISK
var g_requestType = null;
g_requestType = Dialogs.showDialog(new requestTypeDialog(), Constants.DIALOG_TYPE_ACTION, "Update Existing or Upload New Risk?");

try {
    g_oDatabase = ArisData.openDatabase(g_oDatabase.Name(-1), "system", g_strSystemPassword, "dd838074-ac29-11d4-85b8-00005a4053ff", currentLng);

    if (g_requestType != null) {
        var strDialogTitle = "Update Existing Risk";
        if (g_requestType == c_requestTypeNewRisk)
            strDialogTitle = "Upload a New Risk";

        //SHOW MAIN DIALOG
        var dialogResult = Dialogs.showDialog(new mainUserDialog(), Constants.DIALOG_TYPE_WIZARD, strDialogTitle);
        
        if (dialogResult != null && dialogResult.bResult == true) {
            createDocumentAndLogObject(dialogResult.mapSelectedAttributeValues);
        }
    }
}
catch (ex) {
    g_oDatabase.close();
    Dialogs.MsgBox("There was an issue running the report!");
}

function createDocumentAndLogObject(mapDialogValues) {
    var dbVersionLogs = ArisData.openDatabase(c_riskLogDatabaseName, "system", g_strSystemPassword, "dd838074-ac29-11d4-85b8-00005a4053ff", currentLng, false);
    var oLogObject;

    if (dbVersionLogs != null && dbVersionLogs.IsValid()) {
        var oRisk = null;

        mapDialogValues.put(AT_NTT_INITIATOR, oUser.Name(-1));

        if (g_requestType == c_requestTypeExistingRisk) {
            //change existing document
            mapDialogValues.put(AT_NTT_REQUEST_TYPE, "Existing_risk");

            var strRiskGUID = mapDialogValues.get(AT_NTT_RM_RISK_GUID);
            var oRisk = g_oDatabase.FindGUID(strRiskGUID);
        }
        else {
            mapDialogValues.put(AT_NTT_REQUEST_TYPE, "New_risk");
        }

        //Set reorg mark so that the Document is not picked up and reorganized even if it has 0 occurences
        //oDocument.SetReorgMark(1);

        //Create proper folder in the Log object Database
        var listPath = new Array();
        var riskName = mapDialogValues.get(AT_NTT_RM_RISK_NAME)
        listPath.push("Main group");
        listPath.push(riskName);
        var oGroup = dbVersionLogs.Group(listPath, -1);
        if (oGroup == null || !oGroup.IsValid()) {
            oGroup = dbVersionLogs.Group("Main group", -1).CreateChildGroup(riskName, -1);
        }

        //Create new Log Object
        var currentLogCount = dbVersionLogs.RootGroup().ObjDefList(true, [Constants.OT_CASUALOBJ]).length;
        oLogObject = oGroup.CreateObjDef(Constants.OT_CASUALOBJ, riskName, -1);
        oLogObject.Attribute(Constants.AT_NAME, -1).setValue((currentLogCount + 1).toString());

        //Set entered new metadata values to Log Object
        var usedRiskAttNum = new Array()
        var jsonObj = new org.json.JSONObject(); //to store existing atts in JSON

        var hmIterator = mapDialogValues.entrySet().iterator();
        while (hmIterator.hasNext()) {
            var mapElement = hmIterator.next();
            var intAttributeID = parseInt(mapElement.getKey());

            oLogObject.Attribute(intAttributeID, -1).setValue(mapElement.getValue());
            
            usedRiskAttNum.push(mapElement.getKey())
        }

        //create JSON with current risk attributes
        if (g_requestType == c_requestTypeExistingRisk) {
            for (var i = 0; i < usedRiskAttNum.length; i++) {
                var attValue = oRisk.Attribute(usedRiskAttNum[i], -1).getValue();           
                var attName = ArisData.ActiveFilter().AttrTypeName(usedRiskAttNum[i])          
                jsonObj.put(attName, attValue); 
            }
            oLogObject.Attribute(AT_NTT_RM_RISK_ATT_LOG, -1).setValue(jsonObj.toString());
        }

        //Finally start the APG Process
        var apgComponent = Context.getComponent("Process");
        var apgRunResult = apgComponent.run(g_strUpdateRiskAPGGUID, [oLogObject]);

        dbVersionLogs.close();
    }
    else
        Dialogs.MsgBox("Error happened opening the version log database!")

    return oLogObject;
}



///////////////////////////////////////////////////////////////////////////////////////////////////////
// First Dialog to get the the Request Type: New Risk or Update Existing Risk 
///////////////////////////////////////////////////////////////////////////////////////////////////////
function requestTypeDialog() {

    var oResult = null;

    this.getPages = function() {

        var dlgTemplateDocument = Dialogs.createNewDialogTemplate(350, 100, "Please select what do you intend to do?");
        dlgTemplateDocument.GroupBox(15, 20, 320, 100, "Update existing or upload a new Risk?");
        dlgTemplateDocument.OptionGroup("OPT_REQUEST_TYPE");
        dlgTemplateDocument.OptionButton(35, 35, 200, 15, "Update Existing Risk");
        dlgTemplateDocument.OptionButton(35, 60, 200, 15, "Upload a New Risk");
        //dlgTemplateDocument.OptionButton(35, 85, 200, 15, "Archive Document");

        //Check which pages should be showed based on the selected Function Type
        var listPagesToShow = [];
        listPagesToShow.push(dlgTemplateDocument);

        return listPagesToShow;

    }

    this.onClose = function(pageNumber, bOk) {
        if (bOk)
            oResult = this.dialog.getPage(0).getDialogElement("OPT_REQUEST_TYPE").getValue();
    }


    this.getResult = function() {

        return oResult;
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////////
// User Chooser Dialog
///////////////////////////////////////////////////////////////////////////////////////////////////////
function userChooserDialog() {

    var oResults = null

    var userChooserMap = new java.util.TreeMap();

    //Load data for User Selection Table
    var laUsersTableValues = new Array();
    var laSelectedUsersTableValues = new Array();

    var UMC = Context.getComponent("UMC");
    var listAllUsers = ArisData.Unique(UMC.getAllUsers().toArray());
    for (var i = 0; i < listAllUsers.length; i++) {
        var laRowValues = new Array();
        laRowValues.push(listAllUsers[i].getLastName());
        laRowValues.push(listAllUsers[i].getFirstName());
        laRowValues.push(listAllUsers[i].getName());
        laRowValues.push(listAllUsers[i].getEmail());

        laUsersTableValues.push(laRowValues);
    }
    laUsersTableValues.sort();

    var oResult = null;

    this.getPages = function() {

        var dlgTemplateUserChooser = Dialogs.createNewDialogTemplate(800, 500, "Select user");
        dlgTemplateUserChooser.GroupBox(15, 20, 815, 50, "Information");
        dlgTemplateUserChooser.Text(35, 35, 600, 15, "Please select the user");
        //dlgTemplateAddApprowers.Text(35, 50, 600, 15, "NOTE: \"Content Approver\" and \"Final Approver\" must not be the same persons!​ Only one \"Final Approver\" ​​");
        //dlgTemplateAddApprowers.Text(35, 65, 600, 15, "is allowed! In exceptional cases, more than one \"Content Approver\" may be required.");

        dlgTemplateUserChooser.GroupBox(15, 80, 815, 390, "Define Responsibilities:");
        dlgTemplateUserChooser.Text(35, 95, 160, 15, "Search with User Last Name:", "txLblsearchUser");
        dlgTemplateUserChooser.TextBox(195, 95, 220, 15, "User_txtSearch", 0);

        var laColumnHeaders = new Array();
        var laColumnWidths = [75, 75, 50, 100];
        laColumnHeaders.push("Last Name");
        laColumnHeaders.push("First Name");
        laColumnHeaders.push("Approver(s)");
        laColumnHeaders.push("Email");
        dlgTemplateUserChooser.Table(35, 115, 400, 300, laColumnHeaders, null, laColumnWidths, "tblUsers", Constants.TABLE_STYLE_DEFAULT);


        dlgTemplateUserChooser.PushButton(455, 125, 50, 30, "<SYMBOL_ARROWRIGHT>", "UserChooser" + "_btnAdd");
        dlgTemplateUserChooser.PushButton(455, 175, 50, 30, "<SYMBOL_ARROWLEFT>", "UserChooser" + "_btnDelete");


        dlgTemplateUserChooser.Text(525, 95, 300, 15, "Selected User:", "lblSelectedUser");
        dlgTemplateUserChooser.ListBox(525, 115, 285, 100, [], "UserChooser_selection", 0);

        var listPagesToShow = [];
        listPagesToShow.push(dlgTemplateUserChooser);

        return listPagesToShow;

    }

    this.init = function(aPages) {
        aPages[0].getDialogElement("tblUsers").setItems(laUsersTableValues);
    }

    this.onClose = function(pageNumber, bOk) {
        if (bOk)
            //oResult = this.dialog.getDialogElement("UserChooser_selection").getValue();
            oResult = laSelectedUsersTableValues
    }


    this.getResult = function() {

        return oResult;
    }


    //********************************************************************
    //*****  Action Handlers - USER CHOOSER
    //********************************************************************
    this.User_txtSearch_changed = function() {
        this.executeSearch(this.dialog.getPage(0).getDialogElement("User_txtSearch"), laUsersTableValues, this.dialog.getPage(0).getDialogElement("tblUsers"), 0);
    };

    this.UserChooser_btnAdd_pressed = function() {
        this.addUserToSelection(this.dialog.getPage(0).getDialogElement("UserChooser_selection"),
            this.dialog.getPage(0).getDialogElement("tblUsers"),
            userChooserMap,
            laSelectedUsersTableValues);
        if (laSelectedUsersTableValues.length == 1) {
            this.dialog.getPage(0).getDialogElement("UserChooser_btnAdd").setEnabled(false);
            this.dialog.getPage(0).getDialogElement("UserChooser_btnDelete").setEnabled(true);
        }

    };

    this.UserChooser_btnDelete_pressed = function() {
        if (laSelectedUsersTableValues.length == 0)
            return;
        this.deleteUserFromSelection(this.dialog.getPage(0).getDialogElement("UserChooser_selection"),
            this.dialog.getPage(0).getDialogElement("tblUsers"),
            userChooserMap,
            laSelectedUsersTableValues);
        if (laSelectedUsersTableValues.length == 0) {
            this.dialog.getPage(0).getDialogElement("UserChooser_btnAdd").setEnabled(true);
            this.dialog.getPage(0).getDialogElement("UserChooser_btnDelete").setEnabled(false);
        }

    };


    //********************************************************************
    //*****  UTILITY DIALOG FUNCTIONS - USER CHOOSER
    //********************************************************************
    this.executeSearch = function(oSearchBox, listUnfilteredValues, oTableDlgElement, intColumnIndexToSearch) {
        var lsFilterText = oSearchBox.getText();
        var laFilteredTableValues = new Array();

        listUnfilteredValues.forEach(function(row) {
            {
                if (row[intColumnIndexToSearch].toLowerCase().indexOf(lsFilterText.toLowerCase()) !== -1) {
                    laFilteredTableValues.push(row);
                }
            }
        });
        oTableDlgElement.setItems(laFilteredTableValues);
    }

    this.addUserToSelection = function(oSelectionBox, lTable, oSelectedMap, listSelectedUsersArray) {
        var selectedIndexs = lTable.getSelection();
        var tableRows = lTable.getItems();

        if (selectedIndexs != null && selectedIndexs.length > 0) {
            for each(var index in selectedIndexs) {
                var tableRow = tableRows[index];
                if (tableRow != null) {

                    var strKey = tableRow[0].toString() + ", " + tableRow[1].toString() + " (" + tableRow[2].toString() + ")";
                    var strValue = tableRow[2].toString();

                    listSelectedUsersArray.push(tableRow);
                    this.removeFromTable(tableRow[2].toString(), laUsersTableValues);
                    lTable.setItems(laUsersTableValues);

                    oSelectedMap.put(strKey, strValue);
                }
            }

            var listSelectionBoxValues = new Array();
            for each(var oTableRow in listSelectedUsersArray) {
                listSelectionBoxValues.push(oTableRow[0].toString() + ", " + oTableRow[1].toString() + " (" + oTableRow[2].toString() + ")");
            }
            oSelectionBox.setItems(listSelectionBoxValues);

            return true;
        }

        return false;
    }

    this.removeFromTable = function(strId, oTable) {
        var index = 0;
        for each(var oTableRow in oTable) {
            if (oTableRow[2].equals(strId)) {
                oTable.splice(index, 1);
                return;
            }
            index++;
        }
    }

    this.deleteUserFromSelection = function(oSelectionBox, lTable, oSelectedMap, listSelectedUsersArray) {
        var selectedUserIndexs = oSelectionBox.getSelection();
        var l_a_String_listItems = oSelectedMap.keySet();

        if (selectedUserIndexs != null && selectedUserIndexs.length > 0) {
            for each(var index in selectedUserIndexs) {
                var key = l_a_String_listItems.toArray()[index];
                if (key != null) {
                    oSelectedMap.remove(key.toString());
                }

                laUsersTableValues.push(listSelectedUsersArray[index]);
                laUsersTableValues.sort();
                listSelectedUsersArray.splice(index, 1);
                lTable.setItems(laUsersTableValues);

            }
            var listSelectionBoxValues = new Array();
            for each(var oTableRow in listSelectedUsersArray) {
                listSelectionBoxValues.push(oTableRow[0].toString() + ", " + oTableRow[1].toString() + " (" + oTableRow[2].toString() + ")");
            }
            oSelectionBox.setItems(listSelectionBoxValues);

            return true;
        }

        return false;
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////////
// mainUserDialog 
///////////////////////////////////////////////////////////////////////////////////////////////////////
function returnDTO(mapSelectedAttributeValues, bResult) {
    this.mapSelectedAttributeValues = mapSelectedAttributeValues;
    this.testValue1 = "";

    this.bResult = bResult;
}

function mainUserDialog() {

    var objResult;

    var operationalOwnershipMap = new java.util.TreeMap();
    var personNotifiedMap = new java.util.TreeMap();
    var selectedRisksMap = new java.util.TreeMap();

    var pageIndexSelectRisk = 0;
    var pageIndexRiskMetadata = 1;
    var pageIndexRiskMetadata2 = 2;
    var pageIndexRiskMetadata3 = 3;
    var pageIndexRiskMetadata4 = 4;
    var pageIndexAddApprovers = 5; //changed from 3 to 4
    var bMetadataPageFilledInitially = false;


    //Load data for User Selection Table
    var laUsersTableValues = new Array();
    var laSelectedUsersTableValues = new Array();
    var laSelectedPersonNotified = new Array();

    var UMC = Context.getComponent("UMC");
    var listAllUsers = ArisData.Unique(UMC.getAllUsers().toArray());
    for (var i = 0; i < listAllUsers.length; i++) {
        var laRowValues = new Array();
        laRowValues.push(listAllUsers[i].getLastName());
        laRowValues.push(listAllUsers[i].getFirstName());
        laRowValues.push(listAllUsers[i].getName());
        laRowValues.push(listAllUsers[i].getEmail());

        laUsersTableValues.push(laRowValues);
    }
    laUsersTableValues.sort();

    if (g_requestType != c_requestTypeNewRisk) {
        //Load data for Documents Table
        var laRiskTableValues = new Array();

        //GET RISK OBJECTS
        var riskCategoryL0 = g_oDatabase.FindGUID(c_riskRegisterObjectGUID, Constants.CID_OBJDEF)
        var listRiskObjects = new Array()
        if (riskCategoryL0.IsValid()) {
            //get L1 category
            var riskCategoryL1 = riskCategoryL0.getConnectedObjs([Constants.OT_RISK_CATEGORY], Constants.EDGES_OUT, [Constants.CT_BELONG_CAT])
            //get L2 category
            for (var i = 0; i < riskCategoryL1.length; i++) {
                var riskCategoryL2 = riskCategoryL1[i].getConnectedObjs([Constants.OT_RISK_CATEGORY], Constants.EDGES_OUT, [Constants.CT_BELONG_CAT])
                //get risk objects
                for (var j = 0; j < riskCategoryL2.length; j++) {
                    var riskObjects = riskCategoryL2[j].getConnectedObjs([Constants.OT_RISK], Constants.EDGES_OUT, [Constants.CT_SUBS_1])
                    listRiskObjects = listRiskObjects.concat(riskObjects)
                }
            }
        }

        for (var i = 0; i < listRiskObjects.length; i++) {
            var laRowValues = new Array();
            laRowValues.push(listRiskObjects[i].Name(currentLng));
            laRowValues.push(listRiskObjects[i].Attribute(Constants.AT_DESC, currentLng).getValue());
            laRowValues.push(listRiskObjects[i].GUID());

            laRiskTableValues.push(laRowValues);
        }
        laRiskTableValues.sort();
    }


    this.getPages = function() {

        //1.Select Risk Page
        if (g_requestType == c_requestTypeExistingRisk) {
            var dlgTemplateDocument = Dialogs.createNewDialogTemplate(850, 500, "Select Risk to be updated");
            dlgTemplateDocument.GroupBox(15, 20, 815, 65, "Information");
            dlgTemplateDocument.Text(35, 35, 600, 15, "Please select the Risk to be updated by using the search table below.");
            dlgTemplateDocument.Text(35, 50, 600, 15, "Mandatory fields are marked with *.");
            dlgTemplateDocument.Text(35, 65, 600, 15, "A change to the next page is only possible after entering the mandatory information.");

            dlgTemplateDocument.GroupBox(15, 80 + 15, 815, 280 + 15, "Select Risk to be updated:");
            dlgTemplateDocument.Text(35, 95 + 15, 160, 15, "Search with Risk Name:", "txLblsearchDocument");
            dlgTemplateDocument.TextBox(195, 95 + 15, 220, 15, "Risk_txtSearch", 0);

            var laColumnHeaders = new Array();
            var laColumnWidths = [250, 150, 0];
            laColumnHeaders.push("Name");
            laColumnHeaders.push("Description");
            laColumnHeaders.push("GUID");
            dlgTemplateDocument.Table(35, 115 + 15, 400, 235, laColumnHeaders, null, laColumnWidths, "tblRisks", Constants.TABLE_STYLE_DEFAULT);
            
            dlgTemplateDocument.Text(465, 95 + 15, 150, 15, "Select");         
            dlgTemplateDocument.PushButton(455, 115 + 15, 50, 30, "<SYMBOL_ARROWRIGHT>", "Risk" + "_btnAdd");

            dlgTemplateDocument.Text(525, 95 + 15, 150, 15, "Selected Risk: *", "lblSelectedRisk");
            dlgTemplateDocument.ListBox(525, 115 + 15, 285, 30, [], "Risk_selection", 0);

            dlgTemplateDocument.Text(525, 155 + 15, 160, 15, "Risk Description:");
            dlgTemplateDocument.TextBox(525, 170 + 15, 285, 35, "Risk_currentDESC", 1);

            dlgTemplateDocument.Text(525, 210 + 15, 160, 15, "Risk Version:");
            dlgTemplateDocument.TextBox(525, 225 + 15, 285, 15, "Risk_currentVERSION", 0);

            dlgTemplateDocument.Text(525, 250 + 15, 150, 15, "Occurs in Models:", "lblOccInProcesses");
            dlgTemplateDocument.ListBox(525, 265 + 15, 285, 85, [], "Occ_selection", 0);

            dlgTemplateDocument.GroupBox(15, 370 + 30, 815, 80, "Change history");
            dlgTemplateDocument.Text(35, 385 + 30, 600, 15, "Please provide a brief description of the Risk update *");
            dlgTemplateDocument.TextBox(35, 400 + 30, 400, 35, "Risk_newversiondescription", 1);

        }
        else {
            //Don't show the select Document search page in case of NEW Document
            pageIndexSelectRisk = -1;
            pageIndexRiskMetadata = 0;
            pageIndexRiskMetadata2 = 1;
            pageIndexRiskMetadata3 = 2;
            pageIndexRiskMetadata4 = 3;
            pageIndexAddApprovers = 4;
        }

        //2. Change Risk Metadata - Page 1
        var strMetadataTitle = "Risk Basic Information";
        var strMetadataInfo1 = "Please make necessary information update.";
        var strMetadataInfo2 = "Mandatory fields are marked with *.";
        var strMetadataInfo3 = "Continue when ready with “Next”.";
        /*if (g_requestType != c_requestTypeExistingRisk) {
            strMetadataTitle = "Risk Basic Information";
            strMetadataInfo1 = "Please enter the required metadata for the new risk below.";
            strMetadataInfo2 = "After all information is entered, please go forward with 'Next'";
        } */

        var dlgTemplateMetadata = Dialogs.createNewDialogTemplate(850, 500, strMetadataTitle);
        dlgTemplateMetadata.GroupBox(15, 20, 815, 55, "Information");
        dlgTemplateMetadata.Text(35, 25, 600, 15, strMetadataInfo1);
        dlgTemplateMetadata.Text(35, 40, 600, 15, strMetadataInfo2);
        dlgTemplateMetadata.Text(35, 55, 600, 15, strMetadataInfo3);

        dlgTemplateMetadata.GroupBox(15, 80, 815, 390, "Risk Basic Information");
        //dlgTemplateMetadata.CheckBox(35, 95, 780, 15, "Please indicate if the risk metadata must be updated by checking this box.​", "CHKBOX_CHANGE_METADATA", 0);

        dlgTemplateMetadata.Text(35, 105, 140, 15, "Risk Name:*");
        dlgTemplateMetadata.TextBox(195 + 35, 105, 250, 15, "Risk_NAME", 0);

        dlgTemplateMetadata.Text(35, 130, 140, 15, "Risk ID:");
        dlgTemplateMetadata.TextBox(195 + 35, 130, 250, 15, "Risk_ID", 0);

        dlgTemplateMetadata.Text(35, 155, 140, 15, "Status:*");
        dlgTemplateMetadata.ComboBox(195 + 35, 155, 250, 25, listPossibleStatuses, "Risk_STATUS");

        dlgTemplateMetadata.Text(35, 180, 140, 15, "Closing Remarks:");
        dlgTemplateMetadata.TextBox(195 + 35, 180, 250, 40, "Risk_CLOSING_REMARKS", 1);

        dlgTemplateMetadata.Text(35, 225, 140, 15, "Identified:*");
        dlgTemplateMetadata.DateChooser(195 + 35, 225, 250, 15, "Risk_IDENTIFIED")

        dlgTemplateMetadata.Text(35, 250, 140, 15, "Risk identified by:*");
        dlgTemplateMetadata.TextBox(195 + 35, 250, 400, 15, "Risk_IDENTIFIED_BY", 0);

        dlgTemplateMetadata.Text(35, 275, 200, 15, "Executive Management Ownership:*");
        dlgTemplateMetadata.ComboBox(195 + 35, 275, 400, 25, listPossibleExecutiveManager, "Risk_EXECUTIVE_MANAGER");

        dlgTemplateMetadata.Text(35, 300, 140, 15, "Principal Risk:*");
        dlgTemplateMetadata.ComboBox(195 + 35, 300, 400, 25, listPossiblePrincipalRisk, "Risk_PRINCIPAL_RISK");

        dlgTemplateMetadata.Text(35, 325, 140, 15, "Sub Risk:*");
        dlgTemplateMetadata.ComboBox(195 + 35, 325, 400, 25, listPossibleSubRisk, "Risk_SUB_RISK");
        
        dlgTemplateMetadata.Text(35, 350, 140, 15, "Department:");
        dlgTemplateMetadata.ComboBox(195 + 35, 350, 400, 25, listPossibleDepartments, "Risk_DEPARTMENT");

        dlgTemplateMetadata.Text(35, 370+25, 140, 15, "Risk Category:*");
        dlgTemplateMetadata.CheckBox(195 + 35, 350+25, 780, 15, "Strategic Risk", "Risk_STRATEGIC", 2)
        dlgTemplateMetadata.CheckBox(195 + 35, 365+25, 780, 15, "Financial Risk", "Risk_FINANCIAL", 2)
        dlgTemplateMetadata.CheckBox(195 + 35, 380+25, 780, 15, "Operational Risk", "Risk_OPERATIONAL", 2)
        dlgTemplateMetadata.CheckBox(195 + 35, 395+25, 780, 15, "Compliance Risk", "Risk_COMPLIANCE", 2)
        
        //20221020
        //dlgTemplateMetadata.Text(35, 420, 140, 15, "Risk Trend:");
        //dlgTemplateMetadata.ComboBox(195 + 35, 420, 250, 25, listPossibleRiskTrend, "Risk_RISK_TREND"); //TODO MOVE TO PAGE 4

        dlgTemplateMetadata.Text(35, 445, 760, 45, "TXT_METADATA_MESSAGE");

        //3. Change Risk Metadata - Page 2
        var strMetadataTitle2 = "Risk Description";
/*         if (g_requestType != c_requestTypeExistingRisk) {
            strMetadataTitle2 = "Risk Evaluation";
        } */

        var dlgTemplateMetadata2 = Dialogs.createNewDialogTemplate(850, 500, strMetadataTitle2);
        dlgTemplateMetadata2.GroupBox(15, 20, 815, 55, "Information");
        dlgTemplateMetadata2.Text(35, 25, 600, 15, strMetadataInfo1);
        dlgTemplateMetadata2.Text(35, 40, 600, 15, strMetadataInfo2);
        dlgTemplateMetadata2.Text(35, 55, 600, 15, strMetadataInfo3);

        dlgTemplateMetadata2.GroupBox(15, 80, 815, 390, "Risk Description");

        dlgTemplateMetadata2.Text(35, 130, 200, 15, "Risk Description:*");
        dlgTemplateMetadata2.TextBox(195 + 30, 130, 560, 100, "Risk_DESCRIPTION", 1);

        dlgTemplateMetadata2.Text(35, 190 + 50, 200, 15, "Risk Consequence(s):*");
        dlgTemplateMetadata2.TextBox(195 + 30, 190 + 50, 560, 50, "Risk_CONSEQUENCES", 1);

        dlgTemplateMetadata2.Text(35, 260 + 50, 140, 15, "Operational Ownership:*");
        //dlgTemplateMetadata2.ComboBox(195+30, 260, 250, 25, ["SELECT USER -->"], "Risk_OPERATIONAL_OWNERSHIP");
        dlgTemplateMetadata2.TextBox(195 + 30, 260 + 50, 250, 20, "Risk_OPERATIONAL_OWNERSHIP", 0);
        
        if(startEnv != Constants.ENVIRONMENT_TC){    
            dlgTemplateMetadata2.PushButton(455 + 30, 260 + 47, 50, 25, "<SYMBOL_ARROWLEFT>", "OperationalOwnership1" + "_btnAdd");
        }
        else{
            dlgTemplateMetadata2.PushButton(455 + 30, 260 + 50, 50, 20, "🢀", "OperationalOwnership1" + "_btnAdd");
        } 
        
        dlgTemplateMetadata2.Text(455 + 30 + 55, 260 + 52, 140, 15, "Please select user");
        
        dlgTemplateMetadata2.Text(35, 285 + 50, 140, 15, "GDC EMEA Country:*");
        dlgTemplateMetadata2.ComboBox(195 + 30, 285 + 50, 250, 25, listPossibleGDCEMEACountry, "Risk_GDC_EMEA_COUNTRY");

        dlgTemplateMetadata2.Text(35, 310 + 50, 140, 15, "GDC EMEA Affected Site:*");
        dlgTemplateMetadata2.ComboBox(195 + 30, 310 + 50, 250, 25, listPossibleGDCEMEASite, "Risk_GDC_EMEA_SITE");

        dlgTemplateMetadata2.Text(35, 335 + 50, 160, 15, "Corresponding Building(s):");
        dlgTemplateMetadata2.TextBox(195 + 30, 335 + 50, 250, 20, "Risk_CORRESPONDING_BUILDING", 0);

        dlgTemplateMetadata2.Text(35, 360 + 50, 140, 15, "Risk Reach:*");
        dlgTemplateMetadata2.ComboBox(195 + 30, 360 + 50, 560, 25, listPossibleRiskReach, "Risk_REACH");

        dlgTemplateMetadata2.Text(35, 445, 760, 45, "TXT_METADATA_MESSAGE_2");

        //***********************************************************
        //4. Change Risk Metadata - Page 3
        var strMetadataTitle3 = "Risk Evaluation";
/*         if (g_requestType != c_requestTypeExistingRisk) {
            strMetadataTitle3 = "Enter Risk Metadata(3)";
        } */

        var dlgTemplateMetadata3 = Dialogs.createNewDialogTemplate(850, 500, strMetadataTitle3);
        dlgTemplateMetadata3.GroupBox(15, 20, 815, 55, "Information");
        dlgTemplateMetadata3.Text(35, 25, 600, 15, strMetadataInfo1);
        dlgTemplateMetadata3.Text(35, 40, 600, 15, strMetadataInfo2);
        dlgTemplateMetadata3.Text(35, 55, 600, 15, strMetadataInfo3);

        dlgTemplateMetadata3.GroupBox(15, 80, 815, 390, "Risk Evaluation");

        dlgTemplateMetadata3.Text(35, 130, 140, 15, "Risk Impact:*");
        dlgTemplateMetadata3.ComboBox(195 + 30, 130, 250, 25, listPossibleRiskImpact, "Risk_IMPACT");

        dlgTemplateMetadata3.Text(35, 155, 140, 15, "Risk Probability:*");
        dlgTemplateMetadata3.ComboBox(195 + 30, 155, 250, 25, listPossibleRiskProbability, "Risk_PROBABILITY");

        //COLOR BOX INHERENT
        dlgTemplateMetadata3.Text(35, 180, 140, 15, "Gross Risk (Inherent)");
        dlgTemplateMetadata3.Picture(195 + 30, 178, 250, 20, "PICTURE_ORANGE_INHERENT");
        dlgTemplateMetadata3.Picture(195 + 30, 178, 250, 20, "PICTURE_GREEN_INHERENT");
        dlgTemplateMetadata3.Picture(195 + 30, 178, 250, 20, "PICTURE_REDLIGHT_INHERENT");
        dlgTemplateMetadata3.Picture(195 + 30, 178, 250, 20, "PICTURE_REDDARK_INHERENT");

        dlgTemplateMetadata3.Text(35, 205 + 25, 140, 15, "Strength of Controls:*");
        dlgTemplateMetadata3.ComboBox(195 + 30, 205 + 25, 250, 25, listPossibleRiskStrengthOfControls, "Risk_STRENGTH_OF_CONTROLS");

        //COLOR BOX RESIDUAL
        dlgTemplateMetadata3.Text(35, 230 + 25, 140, 15, "Net Risk (Residual)");
        dlgTemplateMetadata3.Picture(195 + 30, 228 + 25, 250, 20, "PICTURE_ORANGE_RESIDUAL");
        dlgTemplateMetadata3.Picture(195 + 30, 228 + 25, 250, 20, "PICTURE_GREEN_RESIDUAL");
        dlgTemplateMetadata3.Picture(195 + 30, 228 + 25, 250, 20, "PICTURE_REDLIGHT_RESIDUAL");
        dlgTemplateMetadata3.Picture(195 + 30, 228 + 25, 250, 20, "PICTURE_REDDARK_RESIDUAL");
        
        dlgTemplateMetadata3.Text(35, 255 + 50, 140, 15, "Risk Trend:*");
        dlgTemplateMetadata3.ComboBox(195 + 30, 255 + 50, 250, 25, listPossibleRiskTrend, "Risk_RISK_TREND"); 

        dlgTemplateMetadata3.Text(35, 445, 760, 45, "TXT_METADATA_MESSAGE_3");

        //***********************************************************
        //4. Change Risk Metadata - Page 4
        var strMetadataTitle4 = "Risk Mitigation";
/*         if (g_requestType != c_requestTypeExistingRisk) {
            strMetadataTitle4 = "Risk Mitigation";
        } */

        var dlgTemplateMetadata4 = Dialogs.createNewDialogTemplate(850, 500, strMetadataTitle4);
        dlgTemplateMetadata4.GroupBox(15, 20, 815, 55, "Information");
        dlgTemplateMetadata4.Text(35, 25, 600, 15, strMetadataInfo1);
        dlgTemplateMetadata4.Text(35, 40, 600, 15, strMetadataInfo2);
        dlgTemplateMetadata4.Text(35, 55, 600, 15, strMetadataInfo3);

        dlgTemplateMetadata4.GroupBox(15, 80, 815, 390, "Risk Mitigation");

        dlgTemplateMetadata4.Text(35, 130, 160, 15, "Treatment Strategy*");
        dlgTemplateMetadata4.ComboBox(195 + 30, 130, 250, 25, listPossibleRiskTreatmentStrategy, "Risk_TREATMENT_STRATEGY");

        dlgTemplateMetadata4.Text(35, 150, 160, 30, "Can the Risk be resolved within 20 working days?*");
        dlgTemplateMetadata4.ComboBox(195 + 30, 155, 250, 25, listPossibleRiskCanBeResolved, "Risk_CAN_BE_RESOLVED");

        dlgTemplateMetadata4.Text(35, 180, 160, 15, "Mitigation Name:");
        dlgTemplateMetadata4.TextBox(195 + 30, 180, 250, 15, "Risk_MITIGATION_NAME", 0);

        dlgTemplateMetadata4.Text(35, 205, 160, 15, "Mitigation Description:");
        dlgTemplateMetadata4.TextBox(195 + 30, 205, 500, 40, "Risk_MITIGATION_DESCRIPTION", 1);

        dlgTemplateMetadata4.Text(35, 250, 160, 15, "Mitigation Responsible:");
        //dlgTemplateMetadata4.TextBox(195+30, 250, 250, 15, "Risk_MITIGATION_RESPONSIBLE",0);    
        dlgTemplateMetadata4.TextBox(195 + 30, 250, 250, 20, "Risk_MITIGATION_RESPONSIBLE", 0);

        if(startEnv != Constants.ENVIRONMENT_TC){    
            dlgTemplateMetadata4.PushButton(455 + 30, 247, 50, 25, "<SYMBOL_ARROWLEFT>", "MitigationResponsible" + "_btnAdd");
        }
        else{
            dlgTemplateMetadata4.PushButton(455 + 30, 250, 50, 20, "🢀", "MitigationResponsible" + "_btnAdd");
        } 
        
        dlgTemplateMetadata4.Text(455 + 30 + 55, 252, 140, 15, "Please select user");
        
        dlgTemplateMetadata4.Text(35, 275, 190, 15, "Estimated Mitigation Costs (EUR €):");
        dlgTemplateMetadata4.TextBox(195 + 30, 275, 250, 15, "Risk_ESTIMATED_MTG_COSTS", 0);
        dlgTemplateMetadata4.Text(195 + 290, 278, 200, 45, "TXT_METADATA_MESSAGE_4_NUM_A");

        dlgTemplateMetadata4.Text(35, 300, 180, 15, "Investment Committee Number:");
        dlgTemplateMetadata4.TextBox(195 + 30, 300, 250, 15, "Risk_INVEST_COM_NUM", 0);
        dlgTemplateMetadata4.Text(195 + 290, 300, 100, 45, "TXT_METADATA_MESSAGE_4_NUM_B");

        dlgTemplateMetadata4.Text(35, 325, 160, 15, "Mitigation Status:");
        dlgTemplateMetadata4.ComboBox(195 + 30, 325, 250, 25, listPossibleRiskMitigationStatus, "Risk_MITIGATION_STATUS");

        dlgTemplateMetadata4.Text(35, 350, 160, 15, "Remarks:");
        dlgTemplateMetadata4.TextBox(195 + 30, 350, 500, 80, "Risk_REMARKS", 1);

        dlgTemplateMetadata4.Text(35, 445, 760, 45, "TXT_METADATA_MESSAGE_4");

        //***********************************************************
        //5. Select Approvers Page
        var dlgTemplateAddApprowers = Dialogs.createNewDialogTemplate(850, 500, "Necessary Notification");
        dlgTemplateAddApprowers.GroupBox(15, 20, 815, 80, "Information");
        dlgTemplateAddApprowers.Text(35, 35, 600, 15, "Please define the “Persons to be notified” for the updated Risk.​");
        dlgTemplateAddApprowers.Text(35, 50, 600, 15, "Submit the updated Risk with “Finish”.");

        dlgTemplateAddApprowers.GroupBox(15, 80 + 30, 815, 390, "");
        dlgTemplateAddApprowers.Text(35, 95 + 30, 160, 15, "Search with User Last Name:", "txLblsearchUser");
        dlgTemplateAddApprowers.TextBox(195, 95 + 30, 220, 15, "User_txtSearch", 0);

        var laColumnHeaders = new Array();
        var laColumnWidths = [75, 75, 50, 100];
        laColumnHeaders.push("Last Name");
        laColumnHeaders.push("First Name");
        laColumnHeaders.push("Approver(s)");
        laColumnHeaders.push("Email");
        dlgTemplateAddApprowers.Table(35, 115 + 30, 400, 335, laColumnHeaders, null, laColumnWidths, "tblUsers", Constants.TABLE_STYLE_DEFAULT);

        dlgTemplateAddApprowers.PushButton(455, 125 + 30, 50, 30, "<SYMBOL_ARROWRIGHT>", "PersonNotified" + "_btnAdd");
        dlgTemplateAddApprowers.PushButton(455, 175 + 30, 50, 30, "<SYMBOL_ARROWLEFT>", "PersonNotified" + "_btnDelete");

        dlgTemplateAddApprowers.Text(525, 95 + 30, 300, 15, "Persons to be notified:*", "lblSelectedPersonNotified");
        dlgTemplateAddApprowers.ListBox(525, 115 + 30, 285, 335, [], "PersonNotified_selection", 0);

        //Check which pages should be showed
        var listPagesToShow = [];

        if (g_requestType == c_requestTypeExistingRisk)
            listPagesToShow.push(dlgTemplateDocument);
        listPagesToShow.push(dlgTemplateMetadata);
        listPagesToShow.push(dlgTemplateMetadata2);
        listPagesToShow.push(dlgTemplateMetadata3);
        listPagesToShow.push(dlgTemplateMetadata4);
        listPagesToShow.push(dlgTemplateAddApprowers);

        return listPagesToShow;
    }

    this.init = function(aPages) {
        var oOrangePicData = Context.getFile("Orange.png", Constants.LOCATION_SCRIPT);
        var oGreenPicData = Context.getFile("Green.png", Constants.LOCATION_SCRIPT);
        var oRedLightPicData = Context.getFile("RedLight.png", Constants.LOCATION_SCRIPT);
        var oRedDarkData = Context.getFile("RedDark.png", Constants.LOCATION_SCRIPT);

        aPages[pageIndexRiskMetadata3].getDialogElement("PICTURE_ORANGE_INHERENT").setPicture(oOrangePicData, "png");
        aPages[pageIndexRiskMetadata3].getDialogElement("PICTURE_GREEN_INHERENT").setPicture(oGreenPicData, "png");
        aPages[pageIndexRiskMetadata3].getDialogElement("PICTURE_REDLIGHT_INHERENT").setPicture(oRedLightPicData, "png");
        aPages[pageIndexRiskMetadata3].getDialogElement("PICTURE_REDDARK_INHERENT").setPicture(oRedDarkData, "png");

        aPages[pageIndexRiskMetadata3].getDialogElement("PICTURE_ORANGE_RESIDUAL").setPicture(oOrangePicData, "png");
        aPages[pageIndexRiskMetadata3].getDialogElement("PICTURE_GREEN_RESIDUAL").setPicture(oGreenPicData, "png");
        aPages[pageIndexRiskMetadata3].getDialogElement("PICTURE_REDLIGHT_RESIDUAL").setPicture(oRedLightPicData, "png");
        aPages[pageIndexRiskMetadata3].getDialogElement("PICTURE_REDDARK_RESIDUAL").setPicture(oRedDarkData, "png");

        aPages[pageIndexRiskMetadata3].getDialogElement("PICTURE_ORANGE_INHERENT").setVisible(false);
        aPages[pageIndexRiskMetadata3].getDialogElement("PICTURE_GREEN_INHERENT").setVisible(false);
        aPages[pageIndexRiskMetadata3].getDialogElement("PICTURE_REDLIGHT_INHERENT").setVisible(false);
        aPages[pageIndexRiskMetadata3].getDialogElement("PICTURE_REDDARK_INHERENT").setVisible(false);

        aPages[pageIndexRiskMetadata3].getDialogElement("PICTURE_ORANGE_RESIDUAL").setVisible(false);
        aPages[pageIndexRiskMetadata3].getDialogElement("PICTURE_GREEN_RESIDUAL").setVisible(false);
        aPages[pageIndexRiskMetadata3].getDialogElement("PICTURE_REDLIGHT_RESIDUAL").setVisible(false);
        aPages[pageIndexRiskMetadata3].getDialogElement("PICTURE_REDDARK_RESIDUAL").setVisible(false);

        aPages[pageIndexRiskMetadata].getDialogElement("Risk_STRATEGIC").setChecked(null);
        aPages[pageIndexRiskMetadata].getDialogElement("Risk_FINANCIAL").setChecked(null);
        aPages[pageIndexRiskMetadata].getDialogElement("Risk_OPERATIONAL").setChecked(null);
        aPages[pageIndexRiskMetadata].getDialogElement("Risk_COMPLIANCE").setChecked(null);

        aPages[pageIndexRiskMetadata].getDialogElement("Risk_ID").setEnabled(false);
        aPages[pageIndexRiskMetadata2].getDialogElement("Risk_OPERATIONAL_OWNERSHIP").setEnabled(false);
        aPages[pageIndexRiskMetadata4].getDialogElement("Risk_MITIGATION_RESPONSIBLE").setEnabled(false);
        
        aPages[pageIndexRiskMetadata4].getDialogElement("TXT_METADATA_MESSAGE_4_NUM_A").setText("");
        aPages[pageIndexRiskMetadata4].getDialogElement("TXT_METADATA_MESSAGE_4_NUM_B").setText("");
        

        if (g_requestType == c_requestTypeExistingRisk) {
            aPages[pageIndexSelectRisk].getDialogElement("tblRisks").setItems(laRiskTableValues);
            aPages[pageIndexSelectRisk].getDialogElement("Occ_selection").setEnabled(false);
            aPages[pageIndexSelectRisk].getDialogElement("Risk_currentDESC").setEnabled(false);
            aPages[pageIndexSelectRisk].getDialogElement("Risk_currentVERSION").setEnabled(false);

            aPages[pageIndexAddApprovers].getDialogElement("tblUsers").setItems(laUsersTableValues);
        }
        else if (g_requestType == c_requestTypeNewRisk) {
            aPages[pageIndexAddApprovers].getDialogElement("tblUsers").setItems(laUsersTableValues);
        }

    }

    //********************************************************************
    //*****  Action Handlers - SELECT RISK
    //********************************************************************
    this.Risk_txtSearch_changed = function() {
        this.executeSearch(this.dialog.getPage(pageIndexSelectRisk).getDialogElement("Risk_txtSearch"), laRiskTableValues, this.dialog.getPage(pageIndexSelectRisk).getDialogElement("tblRisks"), 0);
    };

    this.Risk_btnAdd_pressed = function() {
        selectedRisksMap.clear();
        this.dialog.getPage(pageIndexSelectRisk).getDialogElement("Occ_selection").setItems(new Array());

        if (this.addToSelection(this.dialog.getPage(pageIndexSelectRisk).getDialogElement("Risk_selection"), this.dialog.getPage(pageIndexSelectRisk).getDialogElement("tblRisks"), selectedRisksMap, pageIndexSelectRisk)) {

            try {
                var strRiskGUID = selectedRisksMap.values().toArray()[0];
                var oRisk = g_oDatabase.FindGUID(strRiskGUID);

                var listOccurances = oRisk.OccList();
                var occurancesArray = new Array();
                for each(var oDocOcc in listOccurances) {
                    occurancesArray.push(oDocOcc.Model().Name(currentLng));
                }
                this.dialog.getPage(pageIndexSelectRisk).getDialogElement("Occ_selection").setItems(occurancesArray);
                this.dialog.getPage(pageIndexSelectRisk).getDialogElement("Risk_currentDESC").setText(oRisk.Attribute(Constants.AT_DESC, currentLng).getValue());

                var strMajorVersion = oRisk.Attribute(Constants.AT_MAJOR_VERSION, currentLng).getValue();
                if (strMajorVersion.equals(""))
                    strMajorVersion = "1";
                var strMinorVersion = oRisk.Attribute(Constants.AT_MINOR_VERSION, currentLng).getValue();
                if (strMinorVersion.equals(""))
                    strMinorVersion = "0";
                this.dialog.getPage(pageIndexSelectRisk).getDialogElement("Risk_currentVERSION").setText(strMajorVersion + "." + strMinorVersion);

            }
            catch (ex) {}
        }
    };


    //********************************************************************
    //*****  Action Handlers - RISK METADATA
    //********************************************************************
    this.Risk_IMPACT_selChanged = function(newSelection) {
        var impactValue = listPossibleRiskImpact[newSelection]
        var probabilityValue = listPossibleRiskProbability[this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("Risk_PROBABILITY").getSelectedIndex()]
        var strengthOfControlValue = listPossibleRiskStrengthOfControls[this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("Risk_STRENGTH_OF_CONTROLS").getSelectedIndex()]
        
        if(strengthOfControlValue == "Please select" ){
            var strengthOfControlValue = "";
        }

        var oScores = oRiskRegisterLogic.getRiskScores(impactValue, probabilityValue, strengthOfControlValue);
        try {
            var resultInherentScore = oScores.inherentScore; //should be High
            var resultResidualScore = oScores.residualScore; //should be Medium
        }
        catch (e) {
            var resultInherentScore = ""
            var resultResidualScore = ""
        }

        this.setColorBoxInherent(resultInherentScore)
        this.setColorBoxResidual(resultResidualScore)
    };

    this.Risk_PROBABILITY_selChanged = function(newSelection) {
        var impactValue = listPossibleRiskImpact[this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("Risk_IMPACT").getSelectedIndex()]
        var probabilityValue = listPossibleRiskProbability[newSelection]
        var strengthOfControlValue = listPossibleRiskStrengthOfControls[this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("Risk_STRENGTH_OF_CONTROLS").getSelectedIndex()]
        
        if(strengthOfControlValue == "Please select" ){
            var strengthOfControlValue = "";
        }

        var oScores = oRiskRegisterLogic.getRiskScores(impactValue, probabilityValue, strengthOfControlValue);
        try {
            var resultInherentScore = oScores.inherentScore;
            var resultResidualScore = oScores.residualScore;
        }
        catch (e) {
            var resultInherentScore = ""
            var resultResidualScore = ""
        }

        this.setColorBoxInherent(resultInherentScore)
        this.setColorBoxResidual(resultResidualScore)

    };

    this.Risk_STRENGTH_OF_CONTROLS_selChanged = function(newSelection) {
        var impactValue = listPossibleRiskImpact[this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("Risk_IMPACT").getSelectedIndex()]
        var probabilityValue = listPossibleRiskProbability[this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("Risk_PROBABILITY").getSelectedIndex()]
        var strengthOfControlValue = listPossibleRiskStrengthOfControls[newSelection]
        
        if(strengthOfControlValue == "Please select" ){
            var strengthOfControlValue = "";
        }

        var oScores = oRiskRegisterLogic.getRiskScores(impactValue, probabilityValue, strengthOfControlValue);
        try {
            var resultInherentScore = oScores.inherentScore;
            var resultResidualScore = oScores.residualScore;
        }
        catch (e) {
            var resultInherentScore = ""
            var resultResidualScore = ""
        }
        
        if(strengthOfControlValue.equals("") ){
            var resultResidualScore = "";
        }

        this.setColorBoxInherent(resultInherentScore) 
        this.setColorBoxResidual(resultResidualScore)

    };
    
    this.Risk_ESTIMATED_MTG_COSTS_changed = function(newSelection) {
            //Risk_ESTIMATED_MTG_COSTS  Risk_INVEST_COM_NUM
            var tempValue = this.dialog.getPage(pageIndexRiskMetadata4).getDialogElement("Risk_ESTIMATED_MTG_COSTS").getText()
            if(isNaN(tempValue)){
            this.dialog.getPage(pageIndexRiskMetadata4).getDialogElement("TXT_METADATA_MESSAGE_4_NUM_A").setText("*Number format 123.45");
            
            }   
            else{
                this.dialog.getPage(pageIndexRiskMetadata4).getDialogElement("TXT_METADATA_MESSAGE_4_NUM_A").setText("");
            }
        
    };
    
    this.Risk_INVEST_COM_NUM_changed = function(newSelection) {
            //Risk_ESTIMATED_MTG_COSTS  Risk_INVEST_COM_NUM
            var tempValue = this.dialog.getPage(pageIndexRiskMetadata4).getDialogElement("Risk_INVEST_COM_NUM").getText()
            /*
            if(isNaN(tempValue)){
            this.dialog.getPage(pageIndexRiskMetadata4).getDialogElement("TXT_METADATA_MESSAGE_4_NUM_B").setText("*Value not number");
            
            }   
            else{
                this.dialog.getPage(pageIndexRiskMetadata4).getDialogElement("TXT_METADATA_MESSAGE_4_NUM_B").setText("");
            }
            */
    };

    //********************************************************************
    //*****  Action Handlers - ADD APPROVERS
    //********************************************************************
    this.User_txtSearch_changed = function() {
        this.executeSearch(this.dialog.getPage(pageIndexAddApprovers).getDialogElement("User_txtSearch"), laUsersTableValues, this.dialog.getPage(pageIndexAddApprovers).getDialogElement("tblUsers"), 0);
    };

    this.OperationalOwnership1_btnAdd_pressed = function() {
        this.dialog.setSubDialog("OperationalOwnershipUserDialog", new userChooserDialog(), Constants.DIALOG_TYPE_WIZARD, "Select User") //call first dialog again for test
    };

    this.OperationalOwnershipUserDialog_subDialogClosed = function(subResult, bOk) {
        //subDlg_result = subResult;
        var operationalOwnershipUser = subResult[0][2]
        this.dialog.getPage(pageIndexRiskMetadata2).getDialogElement("Risk_OPERATIONAL_OWNERSHIP").setText(operationalOwnershipUser);
    }

    this.MitigationResponsible_btnAdd_pressed = function() {
        this.dialog.setSubDialog("MitigationResponsibleUserDialog", new userChooserDialog(), Constants.DIALOG_TYPE_WIZARD, "Select User") //call first dialog again for test
    };

    this.MitigationResponsibleUserDialog_subDialogClosed = function(subResult, bOk) {
        //subDlg_result = subResult;
        var mitigationResponsibleUser = subResult[0][2]
        this.dialog.getPage(pageIndexRiskMetadata4).getDialogElement("Risk_MITIGATION_RESPONSIBLE").setText(mitigationResponsibleUser);
    }

    this.OperationalOwnership_btnAdd_pressed = function() {
        this.addUserToSelection(this.dialog.getPage(pageIndexAddApprovers).getDialogElement("OperationalOwnership_selection"),
            this.dialog.getPage(pageIndexAddApprovers).getDialogElement("tblUsers"),
            operationalOwnershipMap,
            laSelectedUsersTableValues);
        if (laSelectedUsersTableValues.length == 1) {
            this.dialog.getPage(pageIndexAddApprovers).getDialogElement("OperationalOwnership_btnAdd").setEnabled(false);
            this.dialog.getPage(pageIndexAddApprovers).getDialogElement("OperationalOwnership_btnDelete").setEnabled(true);
        }

    };

    this.OperationalOwnership_btnDelete_pressed = function() {
        if (laSelectedUsersTableValues.length == 0)
            return;
        this.deleteUserFromSelection(this.dialog.getPage(pageIndexAddApprovers).getDialogElement("OperationalOwnership_selection"),
            this.dialog.getPage(pageIndexAddApprovers).getDialogElement("tblUsers"),
            operationalOwnershipMap,
            laSelectedUsersTableValues);
        if (laSelectedUsersTableValues.length == 0) {
            this.dialog.getPage(pageIndexAddApprovers).getDialogElement("OperationalOwnership_btnAdd").setEnabled(true);
            this.dialog.getPage(pageIndexAddApprovers).getDialogElement("OperationalOwnership_btnDelete").setEnabled(false);
        }

    };

    this.PersonNotified_btnAdd_pressed = function() {
        this.addUserToSelection(this.dialog.getPage(pageIndexAddApprovers).getDialogElement("PersonNotified_selection"),
            this.dialog.getPage(pageIndexAddApprovers).getDialogElement("tblUsers"), personNotifiedMap, laSelectedPersonNotified);
    };

    this.PersonNotified_btnDelete_pressed = function() {
        this.deleteUserFromSelection(this.dialog.getPage(pageIndexAddApprovers).getDialogElement("PersonNotified_selection"),
            this.dialog.getPage(pageIndexAddApprovers).getDialogElement("tblUsers"),
            personNotifiedMap,
            laSelectedPersonNotified);
    };

    //********************************************************************
    //*****  UTILITY DIALOG FUNCTIONS
    //********************************************************************

    this.setColorBoxInherent = function(colorCode) {
        if (colorCode.toLowerCase() == "low") { //green
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_GREEN_INHERENT").setVisible(true);
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_ORANGE_INHERENT").setVisible(false);
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_REDLIGHT_INHERENT").setVisible(false);
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_REDDARK_INHERENT").setVisible(false);
        }
        else if (colorCode.toLowerCase() == "medium") { //orange
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_GREEN_INHERENT").setVisible(false);
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_ORANGE_INHERENT").setVisible(true);
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_REDLIGHT_INHERENT").setVisible(false);
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_REDDARK_INHERENT").setVisible(false);
        }
        else if (colorCode.toLowerCase() == "high") { //red light
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_GREEN_INHERENT").setVisible(false);
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_ORANGE_INHERENT").setVisible(false);
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_REDLIGHT_INHERENT").setVisible(true);
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_REDDARK_INHERENT").setVisible(false);
        }
        else if (colorCode.toLowerCase() == "critical") { //red dark
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_GREEN_INHERENT").setVisible(false);
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_ORANGE_INHERENT").setVisible(false);
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_REDLIGHT_INHERENT").setVisible(false);
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_REDDARK_INHERENT").setVisible(true);
        }
        else { //empty
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_GREEN_INHERENT").setVisible(false);
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_ORANGE_INHERENT").setVisible(false);
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_REDLIGHT_INHERENT").setVisible(false);
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_REDDARK_INHERENT").setVisible(false);
        }
    }

    this.setColorBoxResidual = function(colorCode) {
        if (colorCode.toLowerCase() == "low") { //green
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_GREEN_RESIDUAL").setVisible(true);
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_ORANGE_RESIDUAL").setVisible(false);
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_REDLIGHT_RESIDUAL").setVisible(false);
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_REDDARK_RESIDUAL").setVisible(false);
        }
        else if (colorCode.toLowerCase() == "medium") { //orange
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_GREEN_RESIDUAL").setVisible(false);
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_ORANGE_RESIDUAL").setVisible(true);
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_REDLIGHT_RESIDUAL").setVisible(false);
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_REDDARK_RESIDUAL").setVisible(false);
        }
        else if (colorCode.toLowerCase() == "high") { //red light
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_GREEN_RESIDUAL").setVisible(false);
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_ORANGE_RESIDUAL").setVisible(false);
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_REDLIGHT_RESIDUAL").setVisible(true);
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_REDDARK_RESIDUAL").setVisible(false);
        }
        else if (colorCode.toLowerCase() == "critical") { //red dark
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_GREEN_RESIDUAL").setVisible(false);
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_ORANGE_RESIDUAL").setVisible(false);
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_REDLIGHT_RESIDUAL").setVisible(false);
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_REDDARK_RESIDUAL").setVisible(true);
        }
        else { //empty
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_GREEN_RESIDUAL").setVisible(false);
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_ORANGE_RESIDUAL").setVisible(false);
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_REDLIGHT_RESIDUAL").setVisible(false);
            this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("PICTURE_REDDARK_RESIDUAL").setVisible(false);
        }
    }

    this.executeSearch = function(oSearchBox, listUnfilteredValues, oTableDlgElement, intColumnIndexToSearch) {
        var lsFilterText = oSearchBox.getText();
        var laFilteredTableValues = new Array();

        listUnfilteredValues.forEach(function(row) {
            {
                if (row[intColumnIndexToSearch].toLowerCase().indexOf(lsFilterText.toLowerCase()) !== -1) {
                    laFilteredTableValues.push(row);
                }
            }
        });
        oTableDlgElement.setItems(laFilteredTableValues);
    }

    this.addToSelection = function(oSelectionBox, lTable, oSelectedMap, intPageIndex) {
        var selectedIndexs = lTable.getSelection();
        var tableRows = lTable.getItems();

        if (selectedIndexs != null && selectedIndexs.length > 0) {
            for each(var index in selectedIndexs) {
                var tableRow = tableRows[index];
                if (tableRow != null) {
                    if (intPageIndex == pageIndexSelectRisk) {
                        var strKey = tableRow[0].toString();
                        var strValue = tableRow[2].toString();
                    }
                    oSelectedMap.put(strKey, strValue);
                }
            }

            var listSelectionBoxValues = new Array();
            var itr = oSelectedMap.keySet().iterator();
            while (itr.hasNext()) {
                listSelectionBoxValues.push(itr.next().toString());
            }
            oSelectionBox.setItems(listSelectionBoxValues);

            return true;
        }

        return false;
    }
    this.addUserToSelection = function(oSelectionBox, lTable, oSelectedMap, listSelectedUsersArray) {
        var selectedIndexs = lTable.getSelection();
        var tableRows = lTable.getItems();

        if (selectedIndexs != null && selectedIndexs.length > 0) {
            for each(var index in selectedIndexs) {
                var tableRow = tableRows[index];
                if (tableRow != null) {

                    var strKey = tableRow[0].toString() + ", " + tableRow[1].toString() + " (" + tableRow[2].toString() + ")";
                    var strValue = tableRow[2].toString();

                    listSelectedUsersArray.push(tableRow);
                    this.removeFromTable(tableRow[2].toString(), laUsersTableValues);
                    lTable.setItems(laUsersTableValues);

                    oSelectedMap.put(strKey, strValue);
                }
            }

            var listSelectionBoxValues = new Array();
            for each(var oTableRow in listSelectedUsersArray) {
                listSelectionBoxValues.push(oTableRow[0].toString() + ", " + oTableRow[1].toString() + " (" + oTableRow[2].toString() + ")");
            }
            oSelectionBox.setItems(listSelectionBoxValues);

            return true;
        }

        return false;
    }
    this.removeFromTable = function(strId, oTable) {
        var index = 0;
        for each(var oTableRow in oTable) {
            if (oTableRow[2].equals(strId)) {
                oTable.splice(index, 1);
                return;
            }
            index++;
        }
    }

    this.deleteFromSelection = function(oSelectionBox, oSelectedMap) {
        var selectedUserIndexs = oSelectionBox.getSelection();
        var l_a_String_listItems = oSelectedMap.keySet();

        if (selectedUserIndexs != null && selectedUserIndexs.length > 0) {
            for each(var index in selectedUserIndexs) {
                var key = l_a_String_listItems.toArray()[index];
                if (key != null) {
                    oSelectedMap.remove(key.toString());
                }

            }

            var itemsForSelection = new Array();
            var itr = oSelectedMap.keySet().iterator();
            while (itr.hasNext()) {
                itemsForSelection.push(itr.next());
            }

            oSelectionBox.setItems(itemsForSelection);
        }

    }
    this.deleteUserFromSelection = function(oSelectionBox, lTable, oSelectedMap, listSelectedUsersArray) {
        var selectedUserIndexs = oSelectionBox.getSelection();
        var l_a_String_listItems = oSelectedMap.keySet();

        if (selectedUserIndexs != null && selectedUserIndexs.length > 0) {
            for each(var index in selectedUserIndexs) {
                var key = l_a_String_listItems.toArray()[index];
                if (key != null) {
                    oSelectedMap.remove(key.toString());
                }

                laUsersTableValues.push(listSelectedUsersArray[index]);
                laUsersTableValues.sort();
                listSelectedUsersArray.splice(index, 1);
                lTable.setItems(laUsersTableValues);

            }
            var listSelectionBoxValues = new Array();
            for each(var oTableRow in listSelectedUsersArray) {
                listSelectionBoxValues.push(oTableRow[0].toString() + ", " + oTableRow[1].toString() + " (" + oTableRow[2].toString() + ")");
            }
            oSelectionBox.setItems(listSelectionBoxValues);

            return true;
        }

        return false;
    }

    //********************************************************************
    //*****  DIALOG EVENT HANDLER FUNCTIONS
    //********************************************************************
    this.canGotoNextPage = function(pageNumber) {
        if (pageNumber == pageIndexSelectRisk) {
            if (selectedRisksMap.size() == 0 || this.dialog.getPage(pageIndexSelectRisk).getDialogElement("Risk_newversiondescription").getText().equals(""))
                return false;
            else
                return true;
        }
        else
            return true;
    }
    this.canGotoPreviousPage = function(pageNumber) {
        return true;
    }

    this.onActivatePage = function(pageNumber) {
        if (!bMetadataPageFilledInitially && (pageNumber == pageIndexRiskMetadata || pageNumber == pageIndexRiskMetadata2 || pageNumber == pageIndexRiskMetadata4)) {
            if (g_requestType == c_requestTypeExistingRisk) {
                var strRiskGUID = selectedRisksMap.values().toArray()[0];
                var oRisk = g_oDatabase.FindGUID(strRiskGUID);

                //Page 1 
                this.dialog.getPage(pageIndexRiskMetadata).getDialogElement("Risk_NAME").setText(oRisk.Attribute(Constants.AT_NAME, currentLng).getValue());

                this.dialog.getPage(pageIndexRiskMetadata).getDialogElement("Risk_ID").setText(oRisk.Attribute(AT_NTT_RM_RISK_ID, currentLng).getValue());

                var intValue = getSelectedIndexForValue(oRisk.Attribute(AT_NTT_RM_RISK_STATUS, currentLng).getValue(), listPossibleStatuses);
                this.dialog.getPage(pageIndexRiskMetadata).getDialogElement("Risk_STATUS").setSelection(intValue);

                this.dialog.getPage(pageIndexRiskMetadata).getDialogElement("Risk_CLOSING_REMARKS").setText(oRisk.Attribute(AT_NTT_RM_CLOSING_REMARKS, currentLng).getValue());

                var dateValue = getDateValue(oRisk.Attribute(AT_NTT_RM_RISK_IDENTIFIED_ON, currentLng).getValue())
                this.dialog.getPage(pageIndexRiskMetadata).getDialogElement("Risk_IDENTIFIED").setDate(dateValue);

                this.dialog.getPage(pageIndexRiskMetadata).getDialogElement("Risk_IDENTIFIED_BY").setText(oRisk.Attribute(AT_NTT_RM_RISK_IDENTIFIED_BY, currentLng).getValue());

                var intValue = getSelectedIndexForValue(oRisk.Attribute(AT_NTT_RM_EXEC_MGMT_OWNERSHIP, currentLng).getValue(), listPossibleExecutiveManager);
                this.dialog.getPage(pageIndexRiskMetadata).getDialogElement("Risk_EXECUTIVE_MANAGER").setSelection(intValue);

                var intValue = getSelectedIndexForValue(oRisk.Attribute(AT_NTT_RM_PRINCIPAL_RISK_IMPACTED, currentLng).getValue(), listPossiblePrincipalRisk);
                this.dialog.getPage(pageIndexRiskMetadata).getDialogElement("Risk_PRINCIPAL_RISK").setSelection(intValue);

                var intValue = getSelectedIndexForValue(oRisk.Attribute(AT_NTT_RM_SUB_RISK_CATEGORY, currentLng).getValue(), listPossibleSubRisk);
                this.dialog.getPage(pageIndexRiskMetadata).getDialogElement("Risk_SUB_RISK").setSelection(intValue);
                
                var intValue = getSelectedIndexForValue(oRisk.Attribute(AT_NTT_RM_DEPARTMENT, currentLng).getValue(), listPossibleDepartments);
                this.dialog.getPage(pageIndexRiskMetadata).getDialogElement("Risk_DEPARTMENT").setSelection(intValue);

                var checkBoxValue = getCheckedValue(oRisk.Attribute(AT_NTT_RM_STRATEGIC_RISK, currentLng).getValue())
                this.dialog.getPage(pageIndexRiskMetadata).getDialogElement("Risk_STRATEGIC").setChecked(checkBoxValue);

                var checkBoxValue = getCheckedValue(oRisk.Attribute(AT_NTT_RM_FINANCIAL_RISK, currentLng).getValue())
                this.dialog.getPage(pageIndexRiskMetadata).getDialogElement("Risk_FINANCIAL").setChecked(checkBoxValue);

                var checkBoxValue = getCheckedValue(oRisk.Attribute(AT_NTT_RM_OPERATIONAL_RISK, currentLng).getValue())
                this.dialog.getPage(pageIndexRiskMetadata).getDialogElement("Risk_OPERATIONAL").setChecked(checkBoxValue);

                var checkBoxValue = getCheckedValue(oRisk.Attribute(AT_NTT_RM_COMPLIANCE_RISK, currentLng).getValue())
                this.dialog.getPage(pageIndexRiskMetadata).getDialogElement("Risk_COMPLIANCE").setChecked(checkBoxValue);
                
                //20221020
                //var intValue = getSelectedIndexForValue(oRisk.Attribute(AT_NTT_RM_RISK_TREND, currentLng).getValue(), listPossibleRiskTrend);
                //this.dialog.getPage(pageIndexRiskMetadata).getDialogElement("Risk_RISK_TREND").setSelection(intValue);

                //Page 2
                this.dialog.getPage(pageIndexRiskMetadata2).getDialogElement("Risk_DESCRIPTION").setText(oRisk.Attribute(AT_NTT_RM_RISK_DESCRIPTION, currentLng).getValue());
                this.dialog.getPage(pageIndexRiskMetadata2).getDialogElement("Risk_CONSEQUENCES").setText(oRisk.Attribute(AT_NTT_RM_RISK_CONSEQUENCES, currentLng).getValue());

                //20221115
                this.dialog.getPage(pageIndexRiskMetadata2).getDialogElement("Risk_OPERATIONAL_OWNERSHIP").setText(oRisk.Attribute(AT_NTT_RM_OPERATIONAL_OWNERSHIP, currentLng).getValue());
                
                var intValue = getSelectedIndexForValue(oRisk.Attribute(AT_NTT_RM_GDC_EMEA_AFFECTED_COUNTRY, currentLng).getValue(), listPossibleGDCEMEACountry);
                this.dialog.getPage(pageIndexRiskMetadata2).getDialogElement("Risk_GDC_EMEA_COUNTRY").setSelection(intValue);

                var intValue = getSelectedIndexForValue(oRisk.Attribute(AT_NTT_RM_GDC_EMEA_AFFECTED_SITE, currentLng).getValue(), listPossibleGDCEMEASite);
                this.dialog.getPage(pageIndexRiskMetadata2).getDialogElement("Risk_GDC_EMEA_SITE").setSelection(intValue);

                this.dialog.getPage(pageIndexRiskMetadata2).getDialogElement("Risk_CORRESPONDING_BUILDING").setText(oRisk.Attribute(AT_NTT_RM_AFFECTED_BUILDING, currentLng).getValue());

                var intValue = getSelectedIndexForValue(oRisk.Attribute(AT_NTT_RM_RISK_REACH, currentLng).getValue(), listPossibleRiskReach);
                this.dialog.getPage(pageIndexRiskMetadata2).getDialogElement("Risk_REACH").setSelection(intValue);

                var intValue = getSelectedIndexForValue(oRisk.Attribute(AT_NTT_RM_RISK_REACH, currentLng).getValue(), listPossibleRiskReach);
                this.dialog.getPage(pageIndexRiskMetadata2).getDialogElement("Risk_REACH").setSelection(intValue);


                //Page 3
                var intValue = getSelectedIndexForValue(oRisk.Attribute(AT_NTT_RM_RISK_IMPACT, currentLng).getValue(), listPossibleRiskImpact);
                this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("Risk_IMPACT").setSelection(intValue);

                var intValue = getSelectedIndexForValue(oRisk.Attribute(AT_NTT_RM_RISK_PROBABILITY, currentLng).getValue(), listPossibleRiskProbability);
                this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("Risk_PROBABILITY").setSelection(intValue);

                var intValue = getSelectedIndexForValue(oRisk.Attribute(AT_NTT_RM_STRENGTH_OF_CONTROLS, currentLng).getValue(), listPossibleRiskStrengthOfControls);
                this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("Risk_STRENGTH_OF_CONTROLS").setSelection(intValue);
                
                var intValue = getSelectedIndexForValue(oRisk.Attribute(AT_NTT_RM_RISK_TREND, currentLng).getValue(), listPossibleRiskTrend);
                this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("Risk_RISK_TREND").setSelection(intValue);

                //Page 4
                var intValue = getSelectedIndexForValue(oRisk.Attribute(AT_NTT_RM_RISK_TREATMENT_STRATEGY, currentLng).getValue(), listPossibleRiskTreatmentStrategy);
                this.dialog.getPage(pageIndexRiskMetadata4).getDialogElement("Risk_TREATMENT_STRATEGY").setSelection(intValue);

                var intValue = getSelectedIndexForValue(oRisk.Attribute(AT_NTT_RM_CAN_THIS_BE_FIXED, currentLng).getValue(), listPossibleRiskCanBeResolved);
                this.dialog.getPage(pageIndexRiskMetadata4).getDialogElement("Risk_CAN_BE_RESOLVED").setSelection(intValue);

                this.dialog.getPage(pageIndexRiskMetadata4).getDialogElement("Risk_MITIGATION_NAME").setText(oRisk.Attribute(AT_NTT_RM_MITIGATION_NAME, currentLng).getValue());
                this.dialog.getPage(pageIndexRiskMetadata4).getDialogElement("Risk_MITIGATION_DESCRIPTION").setText(oRisk.Attribute(AT_NTT_RM_MITIGATION_DESCRIPTION, currentLng).getValue());
                this.dialog.getPage(pageIndexRiskMetadata4).getDialogElement("Risk_MITIGATION_RESPONSIBLE").setText(oRisk.Attribute(AT_NTT_RM_MITIGATION_RESPONSIBLE, currentLng).getValue());
                this.dialog.getPage(pageIndexRiskMetadata4).getDialogElement("Risk_ESTIMATED_MTG_COSTS").setText(oRisk.Attribute(AT_NTT_RM_ESTIMATED_MITIGATION_COSTS, currentLng).getValue());
                this.dialog.getPage(pageIndexRiskMetadata4).getDialogElement("Risk_INVEST_COM_NUM").setText(oRisk.Attribute(AT_NTT_RM_INVESTMENT_COMMITTEE_NUMBER, currentLng).getValue());

                var intValue = getSelectedIndexForValue(oRisk.Attribute(AT_NTT_RM_CAN_THIS_BE_FIXED, currentLng).getValue(), listPossibleRiskCanBeResolved);
                this.dialog.getPage(pageIndexRiskMetadata4).getDialogElement("Risk_CAN_BE_RESOLVED").setSelection(intValue);

                this.dialog.getPage(pageIndexRiskMetadata4).getDialogElement("Risk_REMARKS").setText(oRisk.Attribute(AT_NTT_RM_REMARKS, currentLng).getValue()); //DRAFT CHECK CLOSING REMARKS                   


                bMetadataPageFilledInitially = true;
            }
        }

        if (pageNumber == pageIndexRiskMetadata2) {
            var strRiskGUID = selectedRisksMap.values().toArray()[0];
            var oRisk = g_oDatabase.FindGUID(strRiskGUID);

            var impactValue = oRisk.Attribute(AT_NTT_RM_RISK_IMPACT, currentLng).getValue();
            var probabilityValue = oRisk.Attribute(AT_NTT_RM_RISK_PROBABILITY, currentLng).getValue();
            var strengthOfControlValue = oRisk.Attribute(AT_NTT_RM_STRENGTH_OF_CONTROLS, currentLng).getValue();
            //var colorCode = getColorCode(impactValue, probabilityValue)
            
            if(strengthOfControlValue == "Please select" ){
                var strengthOfControlValue = "";
            }

            var oScores = oRiskRegisterLogic.getRiskScores(impactValue, probabilityValue, strengthOfControlValue);
            try {
                var resultInherentScore = oScores.inherentScore; //should be High
                var resultResidualScore = oScores.residualScore; //should be Medium
            }
            catch (e) {
                var resultInherentScore = ""
                var resultResidualScore = ""
            }

            this.setColorBoxInherent(resultInherentScore)
            this.setColorBoxResidual(resultResidualScore)
        }
    }

    this.isInValidState = function(pageNumber) {
        if (pageNumber == pageIndexSelectRisk) {
            if (selectedRisksMap.size() == 0 ||
                this.dialog.getPage(pageNumber).getDialogElement("Risk_newversiondescription").getText().equals(""))
                return false;
            else
                return true;
        }
        else if (pageNumber == pageIndexRiskMetadata) { //PAGE 1

            var selectedPossibleStatus = listPossibleStatuses[this.dialog.getPage(pageNumber).getDialogElement("Risk_STATUS").getValue()];
            var selectedPossibleExecutiveManager = listPossibleExecutiveManager[this.dialog.getPage(pageNumber).getDialogElement("Risk_EXECUTIVE_MANAGER").getValue()];
            var selectedPossiblePrincipalRisk = listPossiblePrincipalRisk[this.dialog.getPage(pageNumber).getDialogElement("Risk_PRINCIPAL_RISK").getValue()];
            var selectedPossibleSubRisk = listPossibleSubRisk[this.dialog.getPage(pageNumber).getDialogElement("Risk_SUB_RISK").getValue()];
            //var selectedPossibleRiskTrend = listPossibleRiskTrend[this.dialog.getPage(pageNumber).getDialogElement("Risk_RISK_TREND").getValue()]; //20201020

            if (selectedPossibleStatus.equals("Please select") || selectedPossibleExecutiveManager.equals("Please select") || selectedPossiblePrincipalRisk.equals("Please select") || selectedPossibleSubRisk.equals("Please select") || 
                this.dialog.getPage(pageNumber).getDialogElement("Risk_STRATEGIC").isUndefined() || this.dialog.getPage(pageNumber).getDialogElement("Risk_FINANCIAL").isUndefined() ||
                this.dialog.getPage(pageNumber).getDialogElement("Risk_COMPLIANCE").isUndefined() || this.dialog.getPage(pageNumber).getDialogElement("Risk_OPERATIONAL").isUndefined() ||
                this.dialog.getPage(pageNumber).getDialogElement("Risk_IDENTIFIED_BY").getText().equals("") || this.dialog.getPage(pageNumber).getDialogElement("Risk_IDENTIFIED").getDate().equals("") ||
                this.dialog.getPage(pageNumber).getDialogElement("Risk_NAME").getText().equals("")) {
                //this.dialog.getPage(pageNumber).getDialogElement("TXT_METADATA_MESSAGE").setText("Please fill in all of the values!");
                this.dialog.getPage(pageNumber).getDialogElement("TXT_METADATA_MESSAGE").setText("");
                return false;
            }
            else {
                this.dialog.getPage(pageNumber).getDialogElement("TXT_METADATA_MESSAGE").setText("");
                return true;
            }
        }
        else if (pageNumber == pageIndexRiskMetadata2) //PAGE 2 
        {
            //var selectedPossibleOperationalOwnership = listPossibleStatuses[this.dialog.getPage(pageNumber).getDialogElement( "Risk_OPERATIONAL_OWNERSHIP" ).getValue()];
            var selectedPossibleGDCEMEACountry = listPossibleGDCEMEACountry[this.dialog.getPage(pageNumber).getDialogElement("Risk_GDC_EMEA_COUNTRY").getValue()];
            var selectedPossibleGDCEMEASite = listPossibleGDCEMEASite[this.dialog.getPage(pageNumber).getDialogElement("Risk_GDC_EMEA_SITE").getValue()];
            var selectedPossibleReach = listPossibleRiskReach[this.dialog.getPage(pageNumber).getDialogElement("Risk_REACH").getValue()];

            if (selectedPossibleGDCEMEACountry.equals("Please select") || selectedPossibleGDCEMEASite.equals("Please select") || selectedPossibleReach.equals("Please select") ||
                this.dialog.getPage(pageNumber).getDialogElement("Risk_DESCRIPTION").getText().equals("") || this.dialog.getPage(pageNumber).getDialogElement("Risk_CONSEQUENCES").getText().equals("") ||
                this.dialog.getPage(pageNumber).getDialogElement("Risk_OPERATIONAL_OWNERSHIP").getText().equals("")) {
                //this.dialog.getPage(pageNumber).getDialogElement("TXT_METADATA_MESSAGE_2").setText("Please fill in all of the values!");
                this.dialog.getPage(pageNumber).getDialogElement("TXT_METADATA_MESSAGE_2").setText("");
                return false;
            }
            else {
                this.dialog.getPage(pageNumber).getDialogElement("TXT_METADATA_MESSAGE_2").setText("");
                return true;
            }

        }
        else if (pageNumber == pageIndexRiskMetadata3) //PAGE 3 
        {

            var selectedPossibleImpact = listPossibleRiskImpact[this.dialog.getPage(pageNumber).getDialogElement("Risk_IMPACT").getValue()];
            var selectedPossibleProbability = listPossibleRiskProbability[this.dialog.getPage(pageNumber).getDialogElement("Risk_PROBABILITY").getValue()];
            var selectedPossibleStrengthOfControls = listPossibleRiskStrengthOfControls[this.dialog.getPage(pageNumber).getDialogElement("Risk_STRENGTH_OF_CONTROLS").getValue()];
            var selectedPossibleRiskTrend = listPossibleRiskTrend[this.dialog.getPage(pageNumber).getDialogElement("Risk_RISK_TREND").getValue()];

            //draft
            if (selectedPossibleImpact.equals("Please select") || selectedPossibleProbability.equals("Please select") || selectedPossibleStrengthOfControls.equals("Please select") ||
                selectedPossibleRiskTrend.equals("Please select") ) {
                //this.dialog.getPage(pageNumber).getDialogElement("TXT_METADATA_MESSAGE_3").setText("Please fill in all of the values!");
                this.dialog.getPage(pageNumber).getDialogElement("TXT_METADATA_MESSAGE_3").setText("");
                return false;
            }
            else {
                this.dialog.getPage(pageNumber).getDialogElement("TXT_METADATA_MESSAGE_3").setText("");
                return true;
            }

        }
        else if (pageNumber == pageIndexRiskMetadata4) //PAGE 4 
        {
            //var selectedPossibleOperationalOwnership = listPossibleStatuses[this.dialog.getPage(pageNumber).getDialogElement( "Risk_OPERATIONAL_OWNERSHIP" ).getValue()];
            var selectedPossibleTreatmentStrategy = listPossibleRiskTreatmentStrategy[this.dialog.getPage(pageNumber).getDialogElement("Risk_TREATMENT_STRATEGY").getValue()];
            var selectedPossibleCanBeResolved = listPossibleRiskCanBeResolved[this.dialog.getPage(pageNumber).getDialogElement("Risk_CAN_BE_RESOLVED").getValue()];
            
            //Risk_ESTIMATED_MTG_COSTS  Risk_INVEST_COM_NUM
            var tempValue = this.dialog.getPage(pageNumber).getDialogElement("Risk_ESTIMATED_MTG_COSTS").getText()
            //var tempValue2 = this.dialog.getPage(pageNumber).getDialogElement("Risk_INVEST_COM_NUM").getText()

            if (selectedPossibleTreatmentStrategy.equals("Please select") || selectedPossibleCanBeResolved.equals("Please select") || isNaN(tempValue))   {
                //this.dialog.getPage(pageNumber).getDialogElement("TXT_METADATA_MESSAGE_4").setText("Please fill in all of the values!");
                this.dialog.getPage(pageNumber).getDialogElement("TXT_METADATA_MESSAGE_4").setText("");
                return false;
            }
            else {
                this.dialog.getPage(pageNumber).getDialogElement("TXT_METADATA_MESSAGE_4").setText("");
                return true;
            }

        }
        else
            return true;
    }

    this.canFinish = function(pageNumber) {
        if (pageNumber == pageIndexAddApprovers) {
            if (laSelectedPersonNotified.length > 0)
                return true;
            else
                return false;
        }
        return false;
    }

    this.onClose = function(pageNumber, bOk) {
        var mapDialogResults = new java.util.HashMap();

        //Existing risk
        if (g_requestType == c_requestTypeExistingRisk) {
            mapDialogResults.put(AT_NTT_RM_RISK_GUID, selectedRisksMap.values().toArray()[0]);
            //mapDialogResults.put(AT_NTT_DATABASE_NAME, g_oDatabase.Name(-1));
            mapDialogResults.put(AT_NTT_RISK_NEW_VERSION_DESCRIPTION, this.dialog.getPage(pageIndexSelectRisk).getDialogElement("Risk_newversiondescription").getText());

        }
        else {
            mapDialogResults.put(AT_NTT_RM_RISK_GUID, "");
            //mapDialogResults.put(AT_NTT_DATABASE_NAME, g_oDatabase.Name(-1));
            mapDialogResults.put(AT_NTT_RISK_NEW_VERSION_DESCRIPTION, "");
        }

        //page 1
        mapDialogResults.put(AT_NTT_RM_RISK_NAME, this.dialog.getPage(pageIndexRiskMetadata).getDialogElement("Risk_NAME").getText());
        mapDialogResults.put(AT_NTT_RM_RISK_ID, this.dialog.getPage(pageIndexRiskMetadata).getDialogElement("Risk_ID").getText());
        mapDialogResults.put(AT_NTT_RM_RISK_STATUS, listPossibleStatuses[this.dialog.getPage(pageIndexRiskMetadata).getDialogElement("Risk_STATUS").getValue()]);
        mapDialogResults.put(AT_NTT_RM_CLOSING_REMARKS, this.dialog.getPage(pageIndexRiskMetadata).getDialogElement("Risk_CLOSING_REMARKS").getText());
        mapDialogResults.put(AT_NTT_RM_RISK_IDENTIFIED_ON, this.dialog.getPage(pageIndexRiskMetadata).getDialogElement("Risk_IDENTIFIED").getDate());
        mapDialogResults.put(AT_NTT_RM_RISK_IDENTIFIED_BY, this.dialog.getPage(pageIndexRiskMetadata).getDialogElement("Risk_IDENTIFIED_BY").getText());
        mapDialogResults.put(AT_NTT_RM_EXEC_MGMT_OWNERSHIP, listPossibleExecutiveManager[this.dialog.getPage(pageIndexRiskMetadata).getDialogElement("Risk_EXECUTIVE_MANAGER").getValue()]);
        mapDialogResults.put(AT_NTT_RM_PRINCIPAL_RISK_IMPACTED, listPossiblePrincipalRisk[this.dialog.getPage(pageIndexRiskMetadata).getDialogElement("Risk_PRINCIPAL_RISK").getValue()]);
        mapDialogResults.put(AT_NTT_RM_SUB_RISK_CATEGORY, listPossibleSubRisk[this.dialog.getPage(pageIndexRiskMetadata).getDialogElement("Risk_SUB_RISK").getValue()]);
        mapDialogResults.put(AT_NTT_RM_DEPARTMENT, listPossibleDepartments[this.dialog.getPage(pageIndexRiskMetadata).getDialogElement("Risk_DEPARTMENT").getValue()]);
        mapDialogResults.put(AT_NTT_RM_STRATEGIC_RISK, this.dialog.getPage(pageIndexRiskMetadata).getDialogElement("Risk_STRATEGIC").isChecked());
        mapDialogResults.put(AT_NTT_RM_FINANCIAL_RISK, this.dialog.getPage(pageIndexRiskMetadata).getDialogElement("Risk_FINANCIAL").isChecked());
        mapDialogResults.put(AT_NTT_RM_OPERATIONAL_RISK, this.dialog.getPage(pageIndexRiskMetadata).getDialogElement("Risk_OPERATIONAL").isChecked());
        mapDialogResults.put(AT_NTT_RM_COMPLIANCE_RISK, this.dialog.getPage(pageIndexRiskMetadata).getDialogElement("Risk_COMPLIANCE").isChecked());
        //mapDialogResults.put(AT_NTT_RM_RISK_TREND, listPossibleRiskTrend[this.dialog.getPage(pageIndexRiskMetadata).getDialogElement("Risk_RISK_TREND").getValue()]); //20201020

        //page 2 
        mapDialogResults.put(AT_NTT_RM_RISK_DESCRIPTION, this.dialog.getPage(pageIndexRiskMetadata2).getDialogElement("Risk_DESCRIPTION").getText());
        mapDialogResults.put(AT_NTT_RM_RISK_CONSEQUENCES, this.dialog.getPage(pageIndexRiskMetadata2).getDialogElement("Risk_CONSEQUENCES").getText());
        mapDialogResults.put(AT_NTT_RM_OPERATIONAL_OWNERSHIP, this.dialog.getPage(pageIndexRiskMetadata2).getDialogElement("Risk_OPERATIONAL_OWNERSHIP").getText());
        mapDialogResults.put(AT_NTT_RM_GDC_EMEA_AFFECTED_COUNTRY, listPossibleGDCEMEACountry[this.dialog.getPage(pageIndexRiskMetadata2).getDialogElement("Risk_GDC_EMEA_COUNTRY").getValue()]);
        mapDialogResults.put(AT_NTT_RM_GDC_EMEA_AFFECTED_SITE, listPossibleGDCEMEASite[this.dialog.getPage(pageIndexRiskMetadata2).getDialogElement("Risk_GDC_EMEA_SITE").getValue()]);
        mapDialogResults.put(AT_NTT_RM_AFFECTED_BUILDING, this.dialog.getPage(pageIndexRiskMetadata2).getDialogElement("Risk_CORRESPONDING_BUILDING").getText());
        mapDialogResults.put(AT_NTT_RM_RISK_REACH, listPossibleRiskReach[this.dialog.getPage(pageIndexRiskMetadata2).getDialogElement("Risk_REACH").getValue()]);

        //page 3
        mapDialogResults.put(AT_NTT_RM_RISK_IMPACT, listPossibleRiskImpact[this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("Risk_IMPACT").getValue()]);
        mapDialogResults.put(AT_NTT_RM_RISK_PROBABILITY, listPossibleRiskProbability[this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("Risk_PROBABILITY").getValue()]);
        mapDialogResults.put(AT_NTT_RM_STRENGTH_OF_CONTROLS, listPossibleRiskStrengthOfControls[this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("Risk_STRENGTH_OF_CONTROLS").getValue()]);
        mapDialogResults.put(AT_NTT_RM_RISK_TREND, listPossibleRiskTrend[this.dialog.getPage(pageIndexRiskMetadata3).getDialogElement("Risk_RISK_TREND").getValue()]);

        //page 4
        mapDialogResults.put(AT_NTT_RM_RISK_TREATMENT_STRATEGY, listPossibleRiskTreatmentStrategy[this.dialog.getPage(pageIndexRiskMetadata4).getDialogElement("Risk_TREATMENT_STRATEGY").getValue()]);
        mapDialogResults.put(AT_NTT_RM_CAN_THIS_BE_FIXED, listPossibleRiskCanBeResolved[this.dialog.getPage(pageIndexRiskMetadata4).getDialogElement("Risk_CAN_BE_RESOLVED").getValue()]);
        mapDialogResults.put(AT_NTT_RM_MITIGATION_NAME, this.dialog.getPage(pageIndexRiskMetadata4).getDialogElement("Risk_MITIGATION_NAME").getText());
        mapDialogResults.put(AT_NTT_RM_MITIGATION_DESCRIPTION, this.dialog.getPage(pageIndexRiskMetadata4).getDialogElement("Risk_MITIGATION_DESCRIPTION").getText());
        mapDialogResults.put(AT_NTT_RM_MITIGATION_RESPONSIBLE, this.dialog.getPage(pageIndexRiskMetadata4).getDialogElement("Risk_MITIGATION_RESPONSIBLE").getText());
        mapDialogResults.put(AT_NTT_RM_ESTIMATED_MITIGATION_COSTS, this.dialog.getPage(pageIndexRiskMetadata4).getDialogElement("Risk_ESTIMATED_MTG_COSTS").getText());
        mapDialogResults.put(AT_NTT_RM_INVESTMENT_COMMITTEE_NUMBER, this.dialog.getPage(pageIndexRiskMetadata4).getDialogElement("Risk_INVEST_COM_NUM").getText());
        mapDialogResults.put(AT_NTT_RM_MITIGATION_STATUS, listPossibleRiskMitigationStatus[this.dialog.getPage(pageIndexRiskMetadata4).getDialogElement("Risk_MITIGATION_STATUS").getValue()]);
        mapDialogResults.put(AT_NTT_RM_REMARKS, this.dialog.getPage(pageIndexRiskMetadata4).getDialogElement("Risk_REMARKS").getText());

        //Add persons to be notified
        var listPersonsToBeNotifiedUsernames = new Array();
        for each(var userRow in laSelectedPersonNotified) {
            listPersonsToBeNotifiedUsernames.push(userRow[2]);
            //listContentApproverNames.push(userRow[0] + " " + userRow[1]);
        }
        mapDialogResults.put(AT_NTT_RISK_PERSONS_TO_BE_NOTIFIED, listPersonsToBeNotifiedUsernames.join(";"));
        //mapDialogResults.put(AT_NTT_DOC_CONTENT_APPROVERNAMES, listContentApproverNames.join("\r\n"));


        if (bOk == true) {
            objResult = new returnDTO(mapDialogResults, true);
        }
        else {
            objResult = new returnDTO(mapDialogResults, false);
        }
    }

    this.getResult = function() {

        return objResult;
    }
}


///////////////////////////////////////////////////////////////////////////////////////////////////////
// HELPER FUNCTIONS 
///////////////////////////////////////////////////////////////////////////////////////////////////////
function getPossibleValues(intAttrType) {
    var retArray = new Array();

    var listValueTypes = g_oFilter.AttrValueTypeNums(intAttrType);
    for each(var valueTypeNum in listValueTypes) {
        retArray.push(g_oFilter.AttrValueType(intAttrType, valueTypeNum));
    }
    if (intAttrType == AT_NTT_RM_SUB_RISK_CATEGORY) {
        retArray = retArray.sort();
    }
      
    if (g_requestType != c_requestTypeExistingRisk)
        retArray.unshift("Please select");
    return retArray;
}

function getSelectedIndexForValue(strValue, listArray) {
    for (var i = 0; i <= listArray.length - 1; i++) {
        if (listArray[i].equals(strValue))
            return i;
    }
    return 0; //changed to 0 form -1
}

function getDateValue(dateStr) {

    formatter = new java.text.SimpleDateFormat("MM/dd/yyyy");
    strdate = new Date(dateStr);
    strgenerationm = formatter.format(strdate);
    return strgenerationm
}

function getCheckedValue(checkStr) {
    if (checkStr.toLowerCase() == "true")
        return true
    else if (checkStr.toLowerCase() == "false")
        return false
    else
        return null
}