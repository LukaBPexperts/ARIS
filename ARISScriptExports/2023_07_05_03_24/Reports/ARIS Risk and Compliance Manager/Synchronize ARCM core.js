/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

/**************************************************
 * Global objects                                 *
 **************************************************/

// dialog y offsets
var gDialogLineY = 0;
const gHEADER = 35;
const gTEXT = 25;
const gINPUT = 15;
const gSAME = 0;
const gSTART = 15;
const gSMALL = 10;
 
// User registry entries
const USER_REGISTRY_SECTION = "ARCM_Synchronization";
const USER_REGISTRY_PATH = "path";
const USER_REGISTRY_PREFIX = "prefix";
const USER_REGISTRY_DISPLAY = "display";
const USER_REGISTRY_SIMULATION_ONLY = "simulationOnly";
const USER_REGISTRY_ACTIVATE_LOGGING = "activateLogging";

// ARCM componenet 
var ARCM = Context.getComponent("ARCM");
// UMC component
var umcComponent = Context.getComponent("UMC");

//Array of ArisObject
var g_aSelectedObjects;

var g_bAllowEntire = false;
var g_bImportModeEntire = false;
var g_bSimulationOnly = false;

var showLogFile = true;
var activateLogging = true;

// Path to log file
var path;

// Prefix of log file
var logFilePrefix = "ARCM_log";

var g_errorMessageProperty = "errorMessage";

var g_bDebug = false;

var isStartedFromConnect = false;
    
/***********************************************
* main function is used for exception handling *
***********************************************/


function synchronizeARCM(isConnect) {   
    isStartedFromConnect = isConnect;
    
    readFromRegistry();
    
    if(!isARCMSyncAdmin() && !ARCM.isARCMAdmin()) {
       Dialogs.MsgBox(formatstring2(getString("PRIVILEGES_MISSING.DBI"), getPrivilegesDisplayName("ARCM_SYNC_ADMIN"), getPrivilegesDisplayName("ARCM_ADMIN")), Constants.MSGBOX_ICON_ERROR, "Error");
        return;            
    }
    if(!isOperateLicense()) {
        Dialogs.MsgBox(formatstring1(getString("OPERATE_PRIVILEGS_MISSING.DBI"),getPrivilegesDisplayName("LICENSE_YRCOP")), Constants.MSGBOX_ICON_ERROR, "Error");
        return;            
    }        

    if(ArisData.getActiveDatabase().IsValid()) {        
        var allClients = ARCM.getAllClients();
        var clientsForDatabase = [];
        
        for (var it = allClients.iterator(); it.hasNext();) {
            var item = it.next();
            if(item.getArisDatabase()==ArisData.getActiveDatabase().Name(0)){
                clientsForDatabase.push(item);
            }
        }    
    }    
    
    g_aSelectedObjects = getSelection();
    
    if(clientsForDatabase.length == 0) {
        Dialogs.MsgBox(getString("NO_DB_FOR_ENVIRONMENT"), Constants.MSGBOX_ICON_INFORMATION, getString("ARCM_SYNC"));
        return;        
    }

    var settings = Dialogs.showDialog(new settingsDialog(clientsForDatabase), Constants.DIALOG_TYPE_ACTION, getString("ARCM_SYNC")); 
                
    if(settings == null) {
        Dialogs.MsgBox(getString("ABORT_SYNC"), Constants.MSGBOX_ICON_INFORMATION, getString("ARCM_SYNC"));
        return;
    }

    try {
        Context.writeStatus(getPropertyPGR("SYNC_DATA_ARCM"));
        var data = startExportReport(settings.client);
    }
    catch (sErrorMessage) {
        var sCompleteMessage = formatstring1(getString("SYNC_REPORT_ERROR"), sErrorMessage);
        Dialogs.MsgBox(sCompleteMessage, Constants.MSGBOX_ICON_INFORMATION, getString("ARCM_SYNC"));
        return;
    }   
    
    var response = ARCM.synchronizeToARCM(settings.client, data, settings.simulationOnly); //new XML import
    
    // Polling every 2s 
    while(response.getExecutionStatus() != Constants.ARCM_JOB_FINSIHED) {
        writeProgressStatus(response);        
        ARCM.sleep(2000);        
        response = ARCM.statusOfJob(response.getId());
    } 
    
    writeProgressStatus(response);

    if(activateLogging == true){
        prepareOutputFile(response);
    }
    
    Dialogs.MsgBox(getResultStatusMessage(response), getResultStatusIcon(response), getString("ARCM_SYNC"));
}

function isOperateLicense() {
    var user = umcComponent.getCurrentUser();
    return findPrivilegesForUser(user, "LICENSE_YRCOP", false, true) != null;  
}

function isARCMSyncAdmin() {
    var user = umcComponent.getCurrentUser();
    return findPrivilegesForUser(user, "ARCM_SYNC_ADMIN", true, false) != null;  
}

function findPrivilegesForUser(user, privilegeName, isFunctionPrivilege, isLicencePrivilege) {
    var user = umcComponent.getCurrentUser();
    var privileges = umcComponent.getPrivilegesForUser(user, isFunctionPrivilege, isLicencePrivilege, true);   
    for(var x = 0; x < privileges.size(); x++) {
        var privilege = privileges.get(x);
        if(privilege.getName() == privilegeName) {
            return privilege;
        }
    }
    return null;    
}

function getPrivilegesDisplayName(privilegeName) {
    var privileges = umcComponent.getAllPrivileges(true, true);   
    for(var x = 0; x < privileges.size(); x++) {
        var privilege = privileges.get(x);
        if(privilege.getName() == privilegeName) {
            return privilege.getDisplayName();
        }
    }
    return null;    
}

function writeProgressStatus(response) {
    var progress = response.getProgress();
    Context.writeStatus(formatstring1(getPropertyPGR("BUILDING_ARCM_STRUCT"), progress) , progress);
}

/*
* Automatically calculate nect row position in settings dialog.
*/
function nextLine(sKind) {
    gDialogLineY += sKind;
    return gDialogLineY;
}

/***********************************************************
* Return the result status icon of the rest call response. *
***********************************************************/

function getResultStatusIcon(response) {
    var status = response.getResultStatus();
    switch(status) {
        case Constants.ARCM_JOB_FAILED:
            return Constants.MSGBOX_ICON_ERROR;
        case Constants.ARCM_JOB_SUCCESS:
            return Constants.MSGBOX_ICON_INFORMATION;
        case Constants.ARCM_JOB_CANCELED:
            return Constants.MSGBOX_ICON_INFORMATION;
        case Constants.ARCM_JOB_WARNINGS:        
            return Constants.MSGBOX_ICON_WARNING;
    }
    return Constants.MSGBOX_ICON_NO;    
}

/**************************************************************
* Return the result status message of the rest call response. *
**************************************************************/

function getResultStatusMessage(response) {
    var withoutLog = "";
    if(activateLogging == false){
        withoutLog = "_WITHOUT_LOG";
    }
    var fileName = Context.getSelectedPath () + java.io.File.separator + Context.getSelectedFile();   
    var status = response.getResultStatus();
    switch(status) {
        case Constants.ARCM_JOB_FAILED:        
            return formatstring1(getPropertyDBI("SYNC_FAILED" + withoutLog), fileName);
        case Constants.ARCM_JOB_SUCCESS:
            return formatstring1(getPropertyDBI("SYNC_SUCCES" + withoutLog), fileName);
        case Constants.ARCM_JOB_CANCELED:
            return formatstring1(getPropertyDBI("SYNC_CANCEL" + withoutLog), fileName);
        case Constants.ARCM_JOB_WARNINGS:        
            return formatstring1(getPropertyDBI("SYNC_WARN" + withoutLog), fileName);
    }
    return formatstring1(getString("SYNC_UNDEF" + withoutLog), fileName);   
}

/*************************************************************
* Starting the export report depending on the ARCM approach. *
*************************************************************/

function startExportReport(client) {
    var reportcomponent = Context.getComponent("Report");
    
    fileName = Context.getSelectedFile();    
    
	var approach = ARCM.getApproach();	
    var scriptID = "f55496e0-055a-11dd-37de-00016ce964d3";
    
	var reportExecData = reportcomponent.createExecInfo(scriptID, g_aSelectedObjects, Context.getSelectedLanguage(), Constants.OutputXML, fileName) ;
	reportExecData.setProperty( "ARCM_APPROACH", approach );
	reportExecData.setProperty( "AUTO_EXECUTION", "true" );
    reportExecData.setProperty( "remoteCalled", "true" );
    //pass the info about the target client to the export report (needed for GUID searches within ARCM database)
    reportExecData.setProperty("ARCM_TARGET_ENVIRONMENT", client.getSign());
    //pass the info if the export report shall create an XML for import "ENTIRE" or not (which means "PARTIAL")
    reportExecData.setProperty("ARCM_IMPORT_MODE_ENTIRE", g_bImportModeEntire.toString());
	    
    var reportResult = reportcomponent.execute(reportExecData);
            
    var aResultFiles = reportResult.getResultFileData();
    
    if (aResultFiles.length > 0) {
        var data = reportResult.getResultFileData()[0].getData();
    } else {
        throw reportResult.getProperty(g_errorMessageProperty);
    }
    
    //in debugging mode: save the transferred XML data as redundant file copy
    if (g_bDebug) {
        try {
            var copyFileName = Context.getSelectedPath() + Context.getSelectedFile();
            var copyFile = new java.io.File(copyFileName);
            copyFile.createNewFile();
            var fos = new java.io.FileOutputStream(copyFileName);
            fos.write(data);
            fos.close();
        } catch (e) {
            //well, do not write a copy then
        }
    }
    
    return data;
}

/*************************************
* Return the selection of the report *
*************************************/

function getSelection() {
    var  selection = ArisData.getSelectedGroups();
    if (selection.length == 0) {
        selection = ArisData.getSelectedModels();
    }
    if (selection.length == 0) {
        selection = ArisData.getSelectedObjDefs();
    }
    if (selection.length == 0) {
        var databases = ArisData.getSelectedDatabases();
        if (databases.length != 0) {
            selection = [databases[0].RootGroup()];
            g_bAllowEntire = true;
        }
    }
    return selection;
}

/*****************************************************************************
* Dialog where all seetings of synchronization are done                      *    
*****************************************************************************/

function settingsDialog(clients) {
        
    var isOk = false;
    var hasEmptyList = true;
    
    // transform the all client data
    var clientNames =  new Array();
    
    if (clients.length == 1){
        clientNames.push(clients[0].getSign() + " (" + clients[0].getName() + ")");
        hasEmptyList = false;
    } else{
        clientNames.push(getString("PLEASE_SELECT"));
        for (var i=0; i<clients.length; i++) {
            var item = clients[i];
            clientNames.push(item.getSign() + " (" + item.getName() + ")");
            hasEmptyList = false;
        }
    }    
  
    // structure of the dialog page
    this.getPages = function() {
        
        var posX = 10;
        var doubleLineLableOffset = 12;
        var textHeight = 16;
		var comboTextHeight = 20;
        var dlgWidth = 460;
        var dlgHeight = 280 + doubleLineLableOffset;
        var lineWidth = dlgWidth - 10;
        var textBoxWidth = dlgWidth - 16;
        var buttonWidth = 40;
        
        gDialogLineY = 0;
        
        var iDialogTemplate = Dialogs.createNewDialogTemplate(dlgWidth, dlgHeight, "First page");
                
        // Layout anpassen,  mit dem inc line pattern aus anderen report               
                
        iDialogTemplate.Text(posX, nextLine(gSTART), lineWidth, textHeight + doubleLineLableOffset, getString("CLIENT_SELECT"));
        iDialogTemplate.ComboBox(posX, nextLine(gINPUT * 2), textBoxWidth, comboTextHeight, clientNames, "CLIENT_CB"); 
        if(!isStartedFromConnect) {                
            iDialogTemplate.CheckBox(posX, nextLine(gTEXT), dlgWidth, textHeight, getString("CHECKBOX_DEACTIVATE_IN_TARGET_ENV"), "ENTIRE_MODE_CHECKBOX", 0);
        }

        iDialogTemplate.CheckBox(posX, nextLine(gHEADER), dlgWidth, textHeight, getString("CHECKBOX_SIMUALTION_ONLY.DBI"), "SIMULATION_ONLY_CHECKBOX", 0);
        
        if(!isStartedFromConnect) {      
            iDialogTemplate.CheckBox(posX, nextLine(gHEADER), textBoxWidth, textHeight, getString("ACTIVATE_LOG"), "ACTIVATE_LOG", 0)                        
            iDialogTemplate.Text(posX, nextLine(gTEXT), lineWidth, textHeight, getString("SELECT_LOG_FOLDER"), "LOG_TEXT");
            iDialogTemplate.TextBox(posX, nextLine(gINPUT), lineWidth - buttonWidth - 10, textHeight, "FOLDER_TXT_EDIT");
            iDialogTemplate.PushButton(lineWidth - buttonWidth + 7, nextLine(gSAME) - 3, buttonWidth, textHeight + 3, "...", "FOLDER_BTN");
            iDialogTemplate.Text(posX, nextLine(gTEXT), lineWidth, textHeight, getString("PREFX_FILE"), "PREFIX_TEXT");
            iDialogTemplate.TextBox(posX, nextLine(gINPUT), textBoxWidth, textHeight, "PREFIX_TXT_EDIT");
            iDialogTemplate.CheckBox(posX, nextLine(gTEXT), textBoxWidth, textHeight, getString("SHOW_LOG"), "DISPLAY_LOG_FILE_CB", 0)                       
        }
        iDialogTemplate.Text(posX, nextLine(gHEADER), lineWidth, textHeight, getString("LOCALE_NOTE"));
        iDialogTemplate.Text(posX, nextLine(gINPUT), lineWidth, textHeight * 3, formatstring1(getString("DB_LANGUAGE"), getDBLanguage(getString("UI_LOCALE_ID"))));
        if(isStartedFromConnect) {   
            iDialogTemplate.HelpButton("HID_5d001470-6e55-11e5-133b-005056c00008_dlg_02.hlp");            
        }
        else {
            iDialogTemplate.HelpButton("HID_5d001470-6e55-11e5-133b-005056c00008_dlg_01.hlp");
        }
             
        return [iDialogTemplate];
    }

    // init of dialog
    this.init = function(aPages) {  
         
        var clientComboBox = this.getDialogElement("CLIENT_CB");
        var comboBoxValues = clientComboBox.getItems();
        if(comboBoxValues.length <2){
            clientComboBox.setEnabled(false);
        }
        if(!isStartedFromConnect) {    		
            //if mode ENTIRE is not allowed then deactivate the mode checkbox
            if (!g_bAllowEntire) {
                var modeCheckbox = this.getDialogElement("ENTIRE_MODE_CHECKBOX");
                modeCheckbox.setEnabled(false);
            }
            var activateLog = this.getDialogElement("ACTIVATE_LOG");
            activateLog.setChecked(activateLogging);
        }		
        var simulationOnly = this.getDialogElement("SIMULATION_ONLY_CHECKBOX");
        simulationOnly.setChecked(g_bSimulationOnly);
                                                
        if(!isStartedFromConnect) {    		
            //use this function also to store the page data locally (for example to access it in "onClose")
            var folderSelectionTextBox = this.getDialogElement("FOLDER_TXT_EDIT");
            folderSelectionTextBox.setText(path);           
            folderSelectionTextBox.setEnabled(false);
            var prefixTextBox = this.getDialogElement("PREFIX_TXT_EDIT");
            prefixTextBox.setText(logFilePrefix);           
            
            var displayLogFileCB = this.getDialogElement("DISPLAY_LOG_FILE_CB");
            displayLogFileCB.setChecked(showLogFile);
        }
    }

    // dilaog will close
    this.onClose = function(pageNumber, bOk) {
        isOk = bOk;
    }

    // check if content of page is valid
    this.isInValidState = function(pageNumber) {
        if(!isStartedFromConnect) {   
            var prefixTextBox = this.dialog.getPage(0).getDialogElement("PREFIX_TXT_EDIT");
            if(!prefixTextBox.getText().matches("[a-zA-Z\\d_\\-]+")) {
                return false;
            }       
            if(activateLogging == true){
                var folderSelectionTextBox = this.getDialogElement("FOLDER_TXT_EDIT");
                if(folderSelectionTextBox.getText() == "") {
                    return false;
                }
            }
        }
        return isValidSelection(this.getDialogElement("CLIENT_CB"));       
    }
    
    function isValidSelection(comboBox){
        if(hasEmptyList){
            return false;
        }
        if(comboBox.getItems().length == 1){
            return true;
        }
        var index = comboBox.getSelection();
        if(index == 0) {
            return false;
        }
        return true;        
    }
    
    // result of the dialog
    this.getResult = function() {
        if(!isOk) {
            return null;
        }
        if(!isStartedFromConnect) {   
            var folderSelectionTextBox = this.getDialogElement("FOLDER_TXT_EDIT");
            path = folderSelectionTextBox.getText();           
            folderSelectionTextBox.setEnabled(false);
            var prefixTextBox = this.getDialogElement("PREFIX_TXT_EDIT");
            logFilePrefix = prefixTextBox.getText();   
    
            var displayLogFileCB = this.getDialogElement("DISPLAY_LOG_FILE_CB");
            showLogFile = displayLogFileCB.isChecked(); 
            
            if (this.getDialogElement("ENTIRE_MODE_CHECKBOX") != null) {
                g_bImportModeEntire = this.getDialogElement("ENTIRE_MODE_CHECKBOX").isChecked();
            }
            
            activateLogging = this.getDialogElement("ACTIVATE_LOG").isChecked();
        }    
        var result = {};
        
        g_bSimulationOnly = this.getDialogElement("SIMULATION_ONLY_CHECKBOX").isChecked();
        result["simulationOnly"] = g_bSimulationOnly;
                
        writeToRegistry();
                
        var index = this.getDialogElement("CLIENT_CB").getSelection();  
        if(clients.length == 1) {
            result["client"] = clients[index];
            return result;
        }
        // First entry in the combox is "Please select" but the clients entry doesn't have this entry. So index is one to high.
        result["client"] = clients[index - 1];
        return result;
    }
    
    // Callback browser folder button "..." is pressed
    this.FOLDER_BTN_pressed = function() {
        this.dialog.setBrowseClientFolders("BrowseFoldersSbdlg", getString("SELECT_LOG_FOLDER"), path, false);  
    }
        
    // Callback Browse sub dialog is closed    
    this.BrowseFoldersSbdlg_subDialogClosed = function(subResult, bOk) {
        if(bOk) {
            this.getDialogElement("FOLDER_TXT_EDIT").setText(subResult);
        }
    }   
    
    // Enable or disable looging components   
    this.ACTIVATE_LOG_selChanged = function() {
        activateLogging = this.getDialogElement("ACTIVATE_LOG").isChecked();
        if(activateLogging){
               this.getDialogElement("LOG_TEXT").setEnabled(true);
               this.getDialogElement("PREFIX_TEXT").setEnabled(true);              
               this.getDialogElement("PREFIX_TXT_EDIT").setEnabled(true);
               this.getDialogElement("DISPLAY_LOG_FILE_CB").setEnabled(true);
               this.getDialogElement("FOLDER_BTN").setEnabled(true);
        }else{
               this.getDialogElement("LOG_TEXT").setEnabled(false);
               this.getDialogElement("PREFIX_TEXT").setEnabled(false); 
               this.getDialogElement("PREFIX_TXT_EDIT").setEnabled(false);
               this.getDialogElement("DISPLAY_LOG_FILE_CB").setEnabled(false);
               this.getDialogElement("FOLDER_BTN").setEnabled(false);
        }    
    }    

    // get dialog element by id
    this.getDialogElement = function(id) {
        return this.dialog.getPage(0).getDialogElement(id);
    }
}

/*************************************************
* Covert a java collection to a javascript array *
*************************************************/

function convertCollectionToJSArray(p_collection) {
    var jsArray = new Array();
    if (p_collection == null) {
        return jsArray;
    } 
    for (var it = p_collection.iterator(); it.hasNext();) {
        jsArray.push(it.next());
    }
    return jsArray;
}

/**************************
* return timestamp format *
**************************/

function getTimeStamp() {
    var today = new java.util.Date();     
    var DATE_FORMAT = new java.text.SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
    return DATE_FORMAT.format(today);
}

/***********************
* return log file name *
***********************/

function getFileName() {
    return logFilePrefix + "-"  + getTimeStamp() + ".txt";
}

/************************************
* Read data from ARIS user registry *
************************************/

function readFromRegistry() {
    if(!isStartedFromConnect) {
        logFilePrefix = Context.getProfileString(USER_REGISTRY_SECTION, USER_REGISTRY_PREFIX, "ARCM_log");
        path = Context.getProfileString(USER_REGISTRY_SECTION, USER_REGISTRY_PATH, "");
        showLogFile = Context.getProfileString(USER_REGISTRY_SECTION, USER_REGISTRY_DISPLAY, "true") == "true";
        g_bSimulationOnly = Context.getProfileString(USER_REGISTRY_SECTION, USER_REGISTRY_SIMULATION_ONLY, "false") == "true";
        activateLogging = Context.getProfileString(USER_REGISTRY_SECTION, USER_REGISTRY_ACTIVATE_LOGGING, "true") == "true";
    }
}

/***********************************
* Write data to ARIS user registry *
***********************************/

function writeToRegistry() {
    if(!isStartedFromConnect) {
        Context.writeProfileString(USER_REGISTRY_SECTION, USER_REGISTRY_PREFIX, logFilePrefix);
        Context.writeProfileString(USER_REGISTRY_SECTION, USER_REGISTRY_PATH, path);
        Context.writeProfileString(USER_REGISTRY_SECTION, USER_REGISTRY_DISPLAY, showLogFile ? "true" : "false");
        Context.writeProfileString(USER_REGISTRY_SECTION, USER_REGISTRY_SIMULATION_ONLY, g_bSimulationOnly ? "true" : "false");    
        Context.writeProfileString(USER_REGISTRY_SECTION, USER_REGISTRY_ACTIVATE_LOGGING, activateLogging ? "true" : "false");    
    }
}

/****************************************************************************
* Prepare output file so that the file is written after report has finished *
****************************************************************************/

function prepareOutputFile(response) {
    Context.setSelectedPath(path);  
    Context.setSelectedFormat(Constants.OUTTEXT);
    Context.setSelectedFile(getFileName());    
    Context.setProperty(Constants.PROPERTY_SHOW_OUTPUT_FILE, showLogFile);
    Context.setProperty(Constants.PROPERTY_SHOW_SUCCESS_MESSAGE, false);       
    
    var oOut = Context.createOutputObject();
    oOut.OutputTxt(response.getResponse());
    oOut.WriteReport();
}

/****************************************************
* Get name of DB language in the language of the OS *
****************************************************/

function getDBLanguage(uiLocaleId) {
    var dbLocaleId = Context.getSelectedLanguage();
    var uiLocale = getLocale(uiLocaleId);
    return getLocale(dbLocaleId).getDisplayName(uiLocale);
}

/******************************
* Get localinfo from local id *
******************************/

function getLocale(localeId) {
  return ArisData.getActiveDatabase().LanguageList()[0].convertLocale(localeId).getLocale();  
}

/****************************************
* Check if exception is an ApiException *
****************************************/

function isApiException(exception) {
    return exception instanceof com.aris.arcm.api.rs.base.exception.ApiException;
}

/***************************************************
* Check if exception is an WebApplicationException *
***************************************************/

function isWebApplicationException(exception) {
    return exception instanceof javax.ws.rs.WebApplicationException;
}

/***********************************************
* Check if exception is an ConnectionException *
***********************************************/

function isConnectionException(exception) {
    return exception instanceof com.aris.arcm.api.rs.base.client.retry.ConnectionException;
}

/*********************************************************
* Check if exception is an ApiPrivilegeMissingException  *
*********************************************************/

function isApiPrivilegeMissingException(exception) {
    return exception instanceof com.aris.arcm.api.rs.base.exception.ApiPrivilegeMissingException;
}

/*********************************
* Handling of exceptions         *
* Write log info to jsreport.log *
*********************************/

function errorHandling(ex) {
    var line = ex.lineNumber;
    var message = ex.message;
    var filename = ex.fileName;
    var exJava = ex.javaException

    if(isApiException(exJava)) {                
        if(isApiPrivilegeMissingException(exJava)) {
            message = formatstring2(getString("PRIVILEGES_MISSING.DBI"), getPrivilegesDisplayName("ARCM_SYNC_ADMIN"), getPrivilegesDisplayName("ARCM_ADMIN"));
        }               
    }
    else if(isWebApplicationException(exJava)) {
        message = getString("SHOW_EXCEPTION");
    }
    else if(isConnectionException(exJava)) {
        message = getString("UNABLE_TO_CONNECT");
    }
    // write log info to jsreport.log
    Context.writeLog("ARCM - Exception in file '" + filename + "', line " + line);
    Context.writeLog("ARCM - Message: '" + ex.message + "'");
 
    if(exJava != null) {
        var aStackTrace = exJava.getStackTrace()
        for(var i = 0; i < aStackTrace.length; i++) {
            Context.writeLog("ARCM -        " + aStackTrace[i].toString());
        }
    }
    
    Dialogs.MsgBox(message, Constants.MSGBOX_ICON_ERROR, "Error");
}

function getPropertyPGR(key) {
    if(g_bSimulationOnly) {
        return getString("SIMULATION_" + key + ".PGR");
    }
    return getString(key);
}

function getPropertyDBI(key) {
    if(g_bSimulationOnly) {
        return getString("SIMULATION_" + key + ".DBI");
    }
    return getString(key);
}

function getPropertyDBT(key) {
    if(g_bSimulationOnly) {
        return getString("SIMULATION_" + key + ".DBT");
    }
    return getString(key);
}

