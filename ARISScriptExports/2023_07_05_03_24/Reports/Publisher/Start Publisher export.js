/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

/*
*   Script to create a Business Publisher export
*   Author: Software AG
*   Date: 08/01/2007
*/




function main()
{
    // Extract Business Publisher script component
    var oBusinessPublisher=Context.getComponent("BusinessPublisher");

    if (oBusinessPublisher==null) {
       Dialogs.MsgBox(getString("BP_1")); 
    }
    else {
        var sBusinessPublisherServerName = "localhost";
        var sExportName = getUniqueExportName();                    // BLUE-7218 Generate unique export names
        var sExportDescription = getString("BP_3");
        
        //Select the active database to create export component instances.
        var databaseToexport = ArisData.getActiveDatabase();
        //Create export component instances.
        var exportData = oBusinessPublisher.createABPExportDataInstance(databaseToexport);
            
        //Check the amount of existing user groups in database
        //If no user group exists, it make so sense to create export
        if (databaseToexport.UserGroupList().length == 0) {
            Dialogs.MsgBox(getString("BP_4"));
            return;
        }

        //set exportname and description
        exportData.setExportName(sExportName);
        exportData.setExportDescription(sExportDescription);
    
        // Specify the web application server address
        exportData.setWebAppServerInfo(exportData.createIBPWebAppServerInfoInstance(sBusinessPublisherServerName));

        // No template
        exportData.setSelectedTemplateAsString("");
        
        //Create language combinations
        var availableInterfaceLanguages = oBusinessPublisher.getAvailableInterfaceLanguages(exportData.getWebAppServerInfo().getServerAddress());
        if (availableInterfaceLanguages.length <= 0) {
            Dialogs.MsgBox(getString("BP_5")); 
            return;
        }

        var selectedLanguagePackages = new Array();
        //Language combinations where interface, database- and methodlanguage are equal
        for ( var i = 0 ; i < availableInterfaceLanguages.length ; i++ ){
            selectedLanguagePackages[i] = exportData.createIBPLanguagePackageDataInstance();
            selectedLanguagePackages[i].setPackageName(getString("BP_6") + i);
            selectedLanguagePackages[i].setPackageDescription(getString("BP_7") + i);
            selectedLanguagePackages[i].setInterfaceLocaleId(availableInterfaceLanguages[i].getLocaleId());
            selectedLanguagePackages[i].setDatabaseLocaleId(availableInterfaceLanguages[i].getLocaleId());
            selectedLanguagePackages[i].setMethodLocaleId(availableInterfaceLanguages[i].getLocaleId());
        }

        exportData.setSelectedLanguagePackages(selectedLanguagePackages);
        //Set first language to default language
        exportData.setDefaultLanguageId(availableInterfaceLanguages[0].getLocaleId());

        // No evaluation filter
        exportData.setSelectedEvalFilterAsString("");

        //All object types allowed
        exportData.setAllowedObjectTypes(databaseToexport.ActiveFilter().ObjTypes());
            
        //Select a specific layout from the existing ones
        var layoutinfos = oBusinessPublisher.getAvailableLayouts(sBusinessPublisherServerName);
        //In this case both export will be exported with the same layout
        exportData.setSelectedLayout(layoutinfos[0].getLayoutName());
       
        //Defaultwert        
        //exportData.setGraphicOptions(AGraphicOptions p_graphicOptions);
        
            //Profile -  Both exports will get the same graphical properties
            var profileData = exportData.getProfileData();     
            profileData.setProfileName(getString("BP_8"));
            profileData.setProfileDescription(getString("BP_9"));
            profileData.setUsePrintScale(false);
            profileData.setInitialScale(100);
            profileData.setScaleValues(new Array(25, 50, 75, 125, 150));
            profileData.setBlackWhite(false);
            profileData.setTransparent(false);
            //Sorting alphabetical
            profileData.setModelSort(0);
            profileData.setObjectSort(0);
            //Sorting methodical
            profileData.setAttributeSort(2);
            profileData.setLinkList(new Array(Constants.AT_LINK, Constants.AT_EXT_1, Constants.AT_EXT_2, Constants.AT_EXT_3));
            profileData.setCopyDocuments(false);
            profileData.setTempFNames(false);
            exportData.setProfileData(profileData);

        //Create export
        if(oBusinessPublisher.exportExists(sBusinessPublisherServerName, sExportName)) {
            // BLUE-7218 Update export if already exists
            oBusinessPublisher.updateExport(databaseToexport, sBusinessPublisherServerName, sExportName);
        } else {
            oBusinessPublisher.createExport(databaseToexport, exportData);
        }        
        //Activate Export
        oBusinessPublisher.activateExport(sBusinessPublisherServerName, exportData.getExportName(), true);
    }
}

function getUniqueExportName() {
    // BLUE-7218 Generate unique export name
    var prefixName = getString("BP_2");

    var dDate = new java.util.Date();
    var format = "yyyyMMdd_HHmmss";
    var dataFormat = new java.text.SimpleDateFormat(format);
    var calendar = java.util.Calendar.getInstance();
    var sDate = dataFormat.format(calendar.getTime());

    return prefixName + "_" + sDate;    
}

main();