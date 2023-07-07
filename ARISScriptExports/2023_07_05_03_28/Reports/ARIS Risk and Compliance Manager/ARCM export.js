/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

/*************************************************************************************************
						 --- Copyright (c) Software AG. All Rights Reserved. ---
*************************************************************************************************/
/* 
Dependencies: 
 - ARCM-Export_ObjectClassification.js
 - ARCM-Util.js
*/

/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
         ~~~~~~~ Global objects, lists, maps ~~~~~~~
  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/

var g_nLoc = Context.getSelectedLanguage();

var g_bDebug = false;

/*
Indicates if the old AT_DEACT flag shall be used to exported objects as deactivated.
Please note that by default object deactivation is not done by this flag anymore but instead 
by executing the report on the ARIS database with option "Delete objects" set to "true".
*/
var g_bReadDeactivateAttribute = false;

// global array of all selected objects
var g_oSelObjDefs = new Array();
// global array of Document IDs
var m_aDocIdList = new Array();                 

// remote execution of report (outside UI by script)
var g_bRemoteCalledProperty = "remoteCalled";
var g_bRemoteCalled = false;

var g_exportRelevantObjectGUIDsProperty = "exportRelevantObjectGUIDs";
var g_errorMessageProperty = "errorMessage";

// for reading and accessing the content of the XML mapping file
var g_xmlConfigReader = null;
// variable for reading parts of the XML mapping file
var xmlConfig_Mapping;

var ARCM = Context.getComponent("ARCM");
var g_arcmEnvironmentID = Context.getProperty("ARCM_TARGET_ENVIRONMENT");
var g_bImportModeEntire = Context.getProperty("ARCM_IMPORT_MODE_ENTIRE");


//---- Cache for Objdef->ARCMAppObject mapping
//format: <ObjDef GUID>|<environmentID> -> <ARCMAppObject GUID>
var hmArisGUID2ArcmGUID = new Packages.java.util.HashMap()
//format: <ObjDef GUID>|<environmentID> -> <ARCMAppObject>
var hmObjDefGUID2ARCMAppObject = new Packages.java.util.HashMap();
//format: <ARCMAppObject> -> Array of <ARCMAppObject>
var hmMasterObject2SubObjects = new Packages.java.util.HashMap();

//---- Cache for generated meta data information
//format: <ARCM ObjType ID>|Node
var hmMeta_ObjTypeID2outputNode = new Packages.java.util.HashMap();
//format: <ARCM ObjType ID>|Set of <String> (the ARCM attribute IDs)
var hmMeta_ObjTypeID2attributeIDs = new Packages.java.util.HashMap();

/*---------------------------------------------------------------------------------------
    Global Maps
---------------------------------------------------------------------------------------*/
var g_objDefMap = new java.util.HashMap();                  // Format: <String>(ObjDef GUID) | <Integer>(export ID)
var g_objDefXMLOutputObjectMap = new java.util.HashMap();   // Format: ObjDef | XMLOutput <object> tag object

/*---------------------------------------------------------------------------------------
    Structure for documents links
---------------------------------------------------------------------------------------*/
function documentID(p_objDef, p_docGUID) {
    this.objDef = p_objDef;
    this.docGUID = p_docGUID;
    return this;
}

/*---------------------------------------------------------------------------------------
    Structure for export info about sections or questions
---------------------------------------------------------------------------------------*/
function exportInfo(p_exportObjDef, p_superiorSectionObjDef, p_aActivatingQuestionObjDefs, p_sArcmGUID) {
    this.exportObjDef = p_exportObjDef;
    this.superiorSectionObjDef = p_superiorSectionObjDef;
    this.aActivatingQuestionObjDefs = p_aActivatingQuestionObjDefs;
    this.arcmGUID = p_sArcmGUID;
    return this;
}

/*---------------------------------------------------------------------------------------
    Structure for output of the object attributes
---------------------------------------------------------------------------------------*/
function attrElement(p_name, p_value) {
    this.name = p_name;
    this.value = p_value;    
    return this;
}

/*---------------------------------------------------------------------------------------
    Structure for output of the object links
---------------------------------------------------------------------------------------*/
function linkElement(p_linkId, p_deact, p_aAttrElements, p_oTargetObjDef, p_sTargetGUID) {  
    this.linkId = p_linkId;
    this.deact = p_deact;
    this.aAttrElements = p_aAttrElements;
    this.targetObjDef = p_oTargetObjDef;
    this.targetGUID = p_sTargetGUID;
    return this;
}

/*---------------------------------------------------------------------------------------
    Structure for sorting info
---------------------------------------------------------------------------------------*/

function sortInfo(p_sSortMethod, p_sSortModelTypeNum) {
    this.sortMethod = p_sSortMethod;
    this.sortModelTypeNum = p_sSortModelTypeNum;
    return this;
}

main();

/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
         ~~~~~~~ Main function & Determination of export mode ~~~~~~~
  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
function main() {

    try {

        //check if the (important for externally passed object GUIDs and error handling)
        var remoteCalledValue = Context.getProperty("remoteCalled");
        if (remoteCalledValue != null && remoteCalledValue.length() > 0 && remoteCalledValue.equalsIgnoreCase( new java.lang.String("true") )) {
            g_bRemoteCalled = true;    
        }
        var bStandAloneExecution = (g_arcmEnvironmentID == null || g_arcmEnvironmentID == "");

        //determine the approach (see arcm-common.js)
        determineApproach(bStandAloneExecution);
        
        //determine the modules to import (see arcm-common.js)
        determineModules();
        
        //read mapping file
        initConfigReader();
    
        //start export generation if config reader is valid
        if (g_xmlConfigReader != null && g_xmlConfigReader.isValid()) {
               
            //I. ----
            // Determine objects to export depending on the selected export mode and start context (group, model, ObjDef)
            // and classify them
            if (g_bTestManagementExport)    {getTestManagementSelectionByContext();}
            if (g_bRiskManagementExport)    {getRiskManagementSelectionByContext();}
            if (g_bSurveyManagementExport)  {getSurveyManagementSelectionByContext();}
            if (g_bAuditManagementExport)   {getAuditManagementSelectionByContext();}
            if (g_bPolicyManagementExport)  {getPolicyManagementSelectionByContext();}
            if (g_bGDPRExport)              {getGDPRSelectionByContext();}
             
            //determine implicit visibility dependencies for survey management 
            if (g_bSurveyManagementExport)  {determineImplicitVisibilityDependencies();}
               
            //II. ---- check hierarchies to export for cycles       
            var sortResultTester = sortTopological(g_testerSet, g_hierarchyParent2hierarchyChildrenHashMap, g_hierarchyChild2hierarchyParentsHashMap);
            var sortResultRegulation = sortTopological(g_regulationSet, g_hierarchyParent2hierarchyChildrenHashMap, g_hierarchyChild2hierarchyParentsHashMap);
            var sortResultOrgUnit = sortTopological(g_orgUnitSet, g_hierarchyParent2hierarchyChildrenHashMap, g_hierarchyChild2hierarchyParentsHashMap);  
            var sortResultProcess = sortTopological(g_processSet, g_hierarchyParent2hierarchyChildrenHashMap, g_hierarchyChild2hierarchyParentsHashMap);
            var sortResultAppSys = sortTopological(g_appSysSet, g_hierarchyParent2hierarchyChildrenHashMap, g_hierarchyChild2hierarchyParentsHashMap);
            var sortResultRiskCategory = sortTopological(g_riskCategorySet, g_hierarchyParent2hierarchyChildrenHashMap, g_hierarchyChild2hierarchyParentsHashMap);
            var sortResultDataCluster = sortTopological(g_dataClusterSet, g_hierarchyParent2hierarchyChildrenHashMap, g_hierarchyChild2hierarchyParentsHashMap);
            
            var hierarchyCyclesArray = new Array();
            hierarchyCyclesArray = hierarchyCyclesArray.concat(sortResultTester.cycleArrays);
            hierarchyCyclesArray = hierarchyCyclesArray.concat(sortResultRegulation.cycleArrays);
            hierarchyCyclesArray = hierarchyCyclesArray.concat(sortResultOrgUnit.cycleArrays);
            hierarchyCyclesArray = hierarchyCyclesArray.concat(sortResultProcess.cycleArrays);
            hierarchyCyclesArray = hierarchyCyclesArray.concat(sortResultAppSys.cycleArrays);
            hierarchyCyclesArray = hierarchyCyclesArray.concat(sortResultRiskCategory.cycleArrays);
            hierarchyCyclesArray = hierarchyCyclesArray.concat(sortResultDataCluster.cycleArrays);
            
            //Abort if there are cycles found   
            if (hierarchyCyclesArray.length > 0) {
                var sErrorMsg = getString("TEXT_7") + "\n";
                for (var i=0; i<hierarchyCyclesArray.length; i++) {
                    var sSubErrorMsg = "- ";
                    for (var j=0; j<hierarchyCyclesArray[i].length; j++) {
                        sSubErrorMsg += "'" + hierarchyCyclesArray[i][j].Name(g_nLoc) + "'";
                        if (j<hierarchyCyclesArray[i].length - 1) {sSubErrorMsg += ", ";}
                    }
                    sErrorMsg += sSubErrorMsg;
                    if (i<hierarchyCyclesArray.length - 1) {sErrorMsg += "\n";}
                }
                if (g_bRemoteCalled) {
                    Context.setProperty(g_errorMessageProperty, sErrorMsg);
                } else {
                    Dialogs.MsgBox(sErrorMsg, Constants.MSGBOX_BTN_OK | Constants.MSGBOX_ICON_ERROR, getString("TEXT_3"));
                }
                return;
            }
            
            g_testerSet = sortResultTester.objDefSet;
            g_regulationSet = sortResultRegulation.objDefSet;
            g_orgUnitSet = sortResultOrgUnit.objDefSet;
            g_processSet = sortResultProcess.objDefSet;
            g_appSysSet = sortResultAppSys.objDefSet;
            g_riskCategorySet = sortResultRiskCategory.objDefSet;
            g_dataClusterSet = sortResultDataCluster.objDefSet;
            
            g_oSelObjDefs = convertHashSetToJSArray(combineGlobalHashSets());
            
            //III. ----
            //Add additional objects to be exported by context property (listed as ";"-separated GUIDs)
            if (g_bRemoteCalled) {
                var exportRelevantObjectGUIDsPropertyValue = Context.getProperty(g_exportRelevantObjectGUIDsProperty);
                
                //var exportRelevantObjectGUIDsPropertyValue = new java.lang.String ( "7f892031-0a53-11de-6fe3-fe72f7be4be1;7fc258a1-0a53-11de-6fe3-fe72f7be4be1;814e28c1-0a53-11de-6fe3-fe72f7be4be1;82b88e31-0a53-11de-6fe3-fe72f7be4be1;806e01f1-0a53-11de-6fe3-fe72f7be4be1;80f385a1-0a53-11de-6fe3-fe72f7be4be1;82309981-0a53-11de-6fe3-fe72f7be4be1;81b6fcb1-0a53-11de-6fe3-fe72f7be4be1;83ff17f1-0a53-11de-6fe3-fe72f7be4be1;835124b1-0a53-11de-6fe3-fe72f7be4be1;80a9ab61-0a53-11de-6fe3-fe72f7be4be1;87cbd761-0a53-11de-6fe3-fe72f7be4be1;86bc1511-0a53-11de-6fe3-fe72f7be4be1;8596cef1-0a53-11de-6fe3-fe72f7be4be1;8650d021-0a53-11de-6fe3-fe72f7be4be1;85e0d041-0a53-11de-6fe3-fe72f7be4be1;8735b1e1-0a53-11de-6fe3-fe72f7be4be1;855d9681-0a53-11de-6fe3-fe72f7be4be1;886b99d1-0a53-11de-6fe3-fe72f7be4be1;e298e301-0a52-11de-6fe3-fe72f7be4be1;e3317981-0a52-11de-6fe3-fe72f7be4be1;dd3e0981-0a52-11de-6fe3-fe72f7be4be1;dc94d131-0a52-11de-6fe3-fe72f7be4be1;db6aa911-0a52-11de-6fe3-fe72f7be4be1;e0154561-0a52-11de-6fe3-fe72f7be4be1;d8f53c41-0a52-11de-6fe3-fe72f7be4be1;d8cf16a1-0a52-11de-6fe3-fe72f7be4be1;d93811a1-0a52-11de-6fe3-fe72f7be4be1;daeec251-0a52-11de-6fe3-fe72f7be4be1;e18b91b1-0a52-11de-6fe3-fe72f7be4be1;dbf9c9b1-0a52-11de-6fe3-fe72f7be4be1;e1015311-0a52-11de-6fe3-fe72f7be4be1;d98b88d1-0a52-11de-6fe3-fe72f7be4be1;de680a91-0a52-11de-6fe3-fe72f7be4be1;df351e91-0a52-11de-6fe3-fe72f7be4be1;d9c4c141-0a52-11de-6fe3-fe72f7be4be1;ddf80ab1-0a52-11de-6fe3-fe72f7be4be1;d7672231-0a52-11de-6fe3-fe72f7be4be1;d72de9c1-0a52-11de-6fe3-fe72f7be4be1;d79473c1-0a52-11de-6fe3-fe72f7be4be1;d21103a1-0a52-11de-6fe3-fe72f7be4be1;d8474901-0a52-11de-6fe3-fe72f7be4be1;cfa77db1-0a52-11de-6fe3-fe72f7be4be1;cd4794b1-0a52-11de-6fe3-fe72f7be4be1;ce84a891-0a52-11de-6fe3-fe72f7be4be1;d04e6c11-0a52-11de-6fe3-fe72f7be4be1;cda481c1-0a52-11de-6fe3-fe72f7be4be1;d0b74001-0a52-11de-6fe3-fe72f7be4be1;ce0d55b1-0a52-11de-6fe3-fe72f7be4be1;cd000461-0a52-11de-6fe3-fe72f7be4be1;cf0c7631-0a52-11de-6fe3-fe72f7be4be1;ccc45af1-0a52-11de-6fe3-fe72f7be4be1;c7c8df81-0a52-11de-6fe3-fe72f7be4be1;c8d630d1-0a52-11de-6fe3-fe72f7be4be1;c88c5691-0a52-11de-6fe3-fe72f7be4be1;c84bf231-0a52-11de-6fe3-fe72f7be4be1;d45463f1-0a52-11de-6fe3-fe72f7be4be1;d5e738f1-0a52-11de-6fe3-fe72f7be4be1;d279d791-0a52-11de-6fe3-fe72f7be4be1;d3043d41-0a52-11de-6fe3-fe72f7be4be1;d4e83f81-0a52-11de-6fe3-fe72f7be4be1;d35eb951-0a52-11de-6fe3-fe72f7be4be1;d39cd3c1-0a52-11de-6fe3-fe72f7be4be1;d2ba3bf1-0a52-11de-6fe3-fe72f7be4be1;d40cd3a1-0a52-11de-6fe3-fe72f7be4be1;d54071a1-0a52-11de-6fe3-fe72f7be4be1;cbb709a1-0a52-11de-6fe3-fe72f7be4be1;ca85dca1-0a52-11de-6fe3-fe72f7be4be1;c9fb76f1-0a52-11de-6fe3-fe72f7be4be1;c9b8c8a1-0a52-11de-6fe3-fe72f7be4be1;cb6abe61-0a52-11de-6fe3-fe72f7be4be1;caf12191-0a52-11de-6fe3-fe72f7be4be1;ca4eee21-0a52-11de-6fe3-fe72f7be4be1;c98de811-0a52-11de-6fe3-fe72f7be4be1;ec679161-0a52-11de-6fe3-fe72f7be4be1;eca313c1-0a52-11de-6fe3-fe72f7be4be1;1b5a4a46-080e-11de-6fe3-fe72f7be4be1;ecef5f01-0a52-11de-6fe3-fe72f7be4be1;ea53cc91-0a52-11de-6fe3-fe72f7be4be1;ea79f231-0a52-11de-6fe3-fe72f7be4be1;eb19b4a1-0a52-11de-6fe3-fe72f7be4be1;eabcc791-0a52-11de-6fe3-fe72f7be4be1;eb803ea1-0a52-11de-6fe3-fe72f7be4be1;e5b7af31-0a52-11de-6fe3-fe72f7be4be1;e69c90f1-0a52-11de-6fe3-fe72f7be4be1;e63ac1e1-0a52-11de-6fe3-fe72f7be4be1;e5f0e7a1-0a52-11de-6fe3-fe72f7be4be1;e7f62d81-0a52-11de-6fe3-fe72f7be4be1;e7731ad1-0a52-11de-6fe3-fe72f7be4be1;e948c531-0a52-11de-6fe3-fe72f7be4be1;e8cf2861-0a52-11de-6fe3-fe72f7be4be1;e863e371-0a52-11de-6fe3-fe72f7be4be1;e7ac5341-0a52-11de-6fe3-fe72f7be4be1;77a609e8-0811-11de-6fe3-fe72f7be4be1;2935f961-0a53-11de-6fe3-fe72f7be4be1;7687ace1-0a53-11de-6fe3-fe72f7be4be1;f36d34b1-0a52-11de-6fe3-fe72f7be4be1;46f2bc91-0a53-11de-6fe3-fe72f7be4be1;1087c471-0a53-11de-6fe3-fe72f7be4be1;f831e8b1-0a52-11de-6fe3-fe72f7be4be1;5f547f31-0a53-11de-6fe3-fe72f7be4be1;51100021-0a53-11de-6fe3-fe72f7be4be1;317b2191-0a53-11de-6fe3-fe72f7be4be1;21789ed1-0a53-11de-6fe3-fe72f7be4be1;35e2e881-0a53-11de-6fe3-fe72f7be4be1;6aaa1bb1-0a53-11de-6fe3-fe72f7be4be1;f14b3f11-0a52-11de-6fe3-fe72f7be4be1;0633b971-0a53-11de-6fe3-fe72f7be4be1;f5265661-0a52-11de-6fe3-fe72f7be4be1;1a4f46e1-0a53-11de-6fe3-fe72f7be4be1;f02d24e1-0a52-11de-6fe3-fe72f7be4be1;581ccf61-0a53-11de-6fe3-fe72f7be4be1;eee1b921-0a52-11de-6fe3-fe72f7be4be1;f1f22d71-0a52-11de-6fe3-fe72f7be4be1;0c99c161-0a53-11de-6fe3-fe72f7be4be1;6e8c5ef1-0a53-11de-6fe3-fe72f7be4be1;2ea3e5b1-0a53-11de-6fe3-fe72f7be4be1;07ef2511-0a53-11de-6fe3-fe72f7be4be1;f44174a1-0a52-11de-6fe3-fe72f7be4be1;f71d6b71-0a52-11de-6fe3-fe72f7be4be1;fa85ead1-0a52-11de-6fe3-fe72f7be4be1;02f138a1-0a53-11de-6fe3-fe72f7be4be1;4098c291-0a53-11de-6fe3-fe72f7be4be1;38ddd901-0a53-11de-6fe3-fe72f7be4be1;fe7b40e1-0a52-11de-6fe3-fe72f7be4be1;6317c8c1-0a53-11de-6fe3-fe72f7be4be1;4a436e31-0a53-11de-6fe3-fe72f7be4be1;54907451-0a53-11de-6fe3-fe72f7be4be1;1296ce51-0a53-11de-6fe3-fe72f7be4be1;048b60a1-0a53-11de-6fe3-fe72f7be4be1;13ab4b91-0a53-11de-6fe3-fe72f7be4be1;f0b4f281-0a52-11de-6fe3-fe72f7be4be1;09c01481-0a53-11de-6fe3-fe72f7be4be1;1f081401-0a53-11de-6fe3-fe72f7be4be1;3bdff571-0a53-11de-6fe3-fe72f7be4be1;ee4922a1-0a52-11de-6fe3-fe72f7be4be1;ef484321-0a52-11de-6fe3-fe72f7be4be1;fd1cc251-0a52-11de-6fe3-fe72f7be4be1;32f16de1-0a53-11de-6fe3-fe72f7be4be1;180be691-0a53-11de-6fe3-fe72f7be4be1;5bb05661-0a53-11de-6fe3-fe72f7be4be1;0e873981-0a53-11de-6fe3-fe72f7be4be1;01658f91-0a53-11de-6fe3-fe72f7be4be1;2695ac01-0a53-11de-6fe3-fe72f7be4be1;4d9b72d1-0a53-11de-6fe3-fe72f7be4be1;f954bdd1-0a52-11de-6fe3-fe72f7be4be1;43be93a1-0a53-11de-6fe3-fe72f7be4be1;0ab80911-0a53-11de-6fe3-fe72f7be4be1;ffe81751-0a52-11de-6fe3-fe72f7be4be1;fbc7e0b1-0a52-11de-6fe3-fe72f7be4be1;f61bd9f1-0a52-11de-6fe3-fe72f7be4be1;66dd5c41-0a53-11de-6fe3-fe72f7be4be1;3d77ac71-0a53-11de-6fe3-fe72f7be4be1;15cfb231-0a53-11de-6fe3-fe72f7be4be1;1ca5ba01-0a53-11de-6fe3-fe72f7be4be1;f2a74ca1-0a52-11de-6fe3-fe72f7be4be1;2be70fa1-0a53-11de-6fe3-fe72f7be4be1;24038f71-0a53-11de-6fe3-fe72f7be4be1;efb38811-0a52-11de-6fe3-fe72f7be4be1;ee825b11-0a52-11de-6fe3-fe72f7be4be1;7288b9e1-0a53-11de-6fe3-fe72f7be4be1;59ea58e1-0a52-11de-6fe3-fe72f7be4be1;5a55c4e1-0a52-11de-6fe3-fe72f7be4be1;5ac835c1-0a52-11de-6fe3-fe72f7be4be1;204432f1-0a52-11de-6fe3-fe72f7be4be1;1dbe2451-0a52-11de-6fe3-fe72f7be4be1;2b3f2c51-0a52-11de-6fe3-fe72f7be4be1;1bbfe351-0a52-11de-6fe3-fe72f7be4be1;258012c1-0a52-11de-6fe3-fe72f7be4be1;231db8c1-0a52-11de-6fe3-fe72f7be4be1;20f70831-0a52-11de-6fe3-fe72f7be4be1;295fe501-0a52-11de-6fe3-fe72f7be4be1;2c8f5301-0a52-11de-6fe3-fe72f7be4be1;248f4a21-0a52-11de-6fe3-fe72f7be4be1;3016f321-0a52-11de-6fe3-fe72f7be4be1;2deb6091-0a52-11de-6fe3-fe72f7be4be1;2780c4c1-0a52-11de-6fe3-fe72f7be4be1;2a95f401-0a52-11de-6fe3-fe72f7be4be1;239029a1-0a52-11de-6fe3-fe72f7be4be1;1beac3e1-0a52-11de-6fe3-fe72f7be4be1;1e296941-0a52-11de-6fe3-fe72f7be4be1;21a2b181-0a52-11de-6fe3-fe72f7be4be1;1ff32cc1-0a52-11de-6fe3-fe72f7be4be1;281e1631-0a52-11de-6fe3-fe72f7be4be1;1eb3a7e1-0a52-11de-6fe3-fe72f7be4be1;2257d0b1-0a52-11de-6fe3-fe72f7be4be1;26780751-0a52-11de-6fe3-fe72f7be4be1;1d6aad21-0a52-11de-6fe3-fe72f7be4be1;2f510b11-0a52-11de-6fe3-fe72f7be4be1;32db1c31-0a52-11de-6fe3-fe72f7be4be1;33276771-0a52-11de-6fe3-fe72f7be4be1;32b03ba1-0a52-11de-6fe3-fe72f7be4be1;328a1601-0a52-11de-6fe3-fe72f7be4be1;335be4f1-0a52-11de-6fe3-fe72f7be4be1;565487f1-0a52-11de-6fe3-fe72f7be4be1;53147821-0a52-11de-6fe3-fe72f7be4be1;54374d41-0a52-11de-6fe3-fe72f7be4be1;5728a0d1-0a52-11de-6fe3-fe72f7be4be1;518d62f1-0a52-11de-6fe3-fe72f7be4be1;558e9fe1-0a52-11de-6fe3-fe72f7be4be1;522f9661-0a52-11de-6fe3-fe72f7be4be1;580d8291-0a52-11de-6fe3-fe72f7be4be1;51542a81-0a52-11de-6fe3-fe72f7be4be1;539eb6c1-0a52-11de-6fe3-fe72f7be4be1;51d76441-0a52-11de-6fe3-fe72f7be4be1;529adb51-0a52-11de-6fe3-fe72f7be4be1;54dbcaa1-0a52-11de-6fe3-fe72f7be4be1;46488f51-0a52-11de-6fe3-fe72f7be4be1;46af1951-0a52-11de-6fe3-fe72f7be4be1;47750161-0a52-11de-6fe3-fe72f7be4be1;47cd3381-0a52-11de-6fe3-fe72f7be4be1;46e5e0c1-0a52-11de-6fe3-fe72f7be4be1;43087f81-0a52-11de-6fe3-fe72f7be4be1;441841d1-0a52-11de-6fe3-fe72f7be4be1;44990a91-0a52-11de-6fe3-fe72f7be4be1;452f3011-0a52-11de-6fe3-fe72f7be4be1;43a11601-0a52-11de-6fe3-fe72f7be4be1;42be7e31-0a52-11de-6fe3-fe72f7be4be1;433f46f1-0a52-11de-6fe3-fe72f7be4be1;428545c1-0a52-11de-6fe3-fe72f7be4be1;3eb3cb61-0a52-11de-6fe3-fe72f7be4be1;40e418e1-0a52-11de-6fe3-fe72f7be4be1;3f6b5b91-0a52-11de-6fe3-fe72f7be4be1;3e4d4161-0a52-11de-6fe3-fe72f7be4be1;3dfc3b31-0a52-11de-6fe3-fe72f7be4be1;3ff35041-0a52-11de-6fe3-fe72f7be4be1;414376f1-0a52-11de-6fe3-fe72f7be4be1;3d982231-0a52-11de-6fe3-fe72f7be4be1;3ef1bec1-0a52-11de-6fe3-fe72f7be4be1;40445671-0a52-11de-6fe3-fe72f7be4be1;3dbbd6d1-0a52-11de-6fe3-fe72f7be4be1;385c1b51-0a52-11de-6fe3-fe72f7be4be1;3ae25101-0a52-11de-6fe3-fe72f7be4be1;34137521-0a52-11de-6fe3-fe72f7be4be1;34f12af1-0a52-11de-6fe3-fe72f7be4be1;3559fee1-0a52-11de-6fe3-fe72f7be4be1;344cad91-0a52-11de-6fe3-fe72f7be4be1;39fd6f41-0a52-11de-6fe3-fe72f7be4be1;37479e11-0a52-11de-6fe3-fe72f7be4be1;35d151c1-0a52-11de-6fe3-fe72f7be4be1;361d9d01-0a52-11de-6fe3-fe72f7be4be1;37f59151-0a52-11de-6fe3-fe72f7be4be1;36ac9691-0a52-11de-6fe3-fe72f7be4be1;3926be51-0a52-11de-6fe3-fe72f7be4be1;3bda4591-0a52-11de-6fe3-fe72f7be4be1;3498f8d1-0a52-11de-6fe3-fe72f7be4be1;a9aa4d41-0a52-11de-6fe3-fe72f7be4be1;c0b27351-0a52-11de-6fe3-fe72f7be4be1;ab1268c1-0a52-11de-6fe3-fe72f7be4be1;b7813d71-0a52-11de-6fe3-fe72f7be4be1;a9521b21-0a52-11de-6fe3-fe72f7be4be1;b63387c1-0a52-11de-6fe3-fe72f7be4be1;adccf4e1-0a52-11de-6fe3-fe72f7be4be1;b0878101-0a52-11de-6fe3-fe72f7be4be1;ba53be61-0a52-11de-6fe3-fe72f7be4be1;c4413f61-0a52-11de-6fe3-fe72f7be4be1;ad0251e1-0a52-11de-6fe3-fe72f7be4be1;bbd83b81-0a52-11de-6fe3-fe72f7be4be1;aa8a7411-0a52-11de-6fe3-fe72f7be4be1;c2777be1-0a52-11de-6fe3-fe72f7be4be1;ac4d0ba1-0a52-11de-6fe3-fe72f7be4be1;aeaf65a1-0a52-11de-6fe3-fe72f7be4be1;b4f402e1-0a52-11de-6fe3-fe72f7be4be1;b1901761-0a52-11de-6fe3-fe72f7be4be1;bd68c691-0a52-11de-6fe3-fe72f7be4be1;a8235f21-0a52-11de-6fe3-fe72f7be4be1;aa134841-0a52-11de-6fe3-fe72f7be4be1;b2a000c1-0a52-11de-6fe3-fe72f7be4be1;bf055f91-0a52-11de-6fe3-fe72f7be4be1;af91fd71-0a52-11de-6fe3-fe72f7be4be1;a7ea26b1-0a52-11de-6fe3-fe72f7be4be1;b8e476f1-0a52-11de-6fe3-fe72f7be4be1;b3c2d5e1-0a52-11de-6fe3-fe72f7be4be1;aba88e41-0a52-11de-6fe3-fe72f7be4be1;a3446c61-0a52-11de-6fe3-fe72f7be4be1;a5f311a1-0a52-11de-6fe3-fe72f7be4be1;a69a0001-0a52-11de-6fe3-fe72f7be4be1;a4d28671-0a52-11de-6fe3-fe72f7be4be1;a308c2f1-0a52-11de-6fe3-fe72f7be4be1;a38e46a1-0a52-11de-6fe3-fe72f7be4be1;a3e8e9c1-0a52-11de-6fe3-fe72f7be4be1;a5580a21-0a52-11de-6fe3-fe72f7be4be1;a4569fb1-0a52-11de-6fe3-fe72f7be4be1;8cd4dcd1-0a52-11de-6fe3-fe72f7be4be1;8ac843f1-0a52-11de-6fe3-fe72f7be4be1;88ad7a41-0a52-11de-6fe3-fe72f7be4be1;9addd981-0a52-11de-6fe3-fe72f7be4be1;a0235641-0a52-11de-6fe3-fe72f7be4be1;9426cb61-0a52-11de-6fe3-fe72f7be4be1;8fd6f941-0a52-11de-6fe3-fe72f7be4be1;8d9ac4e1-0a52-11de-6fe3-fe72f7be4be1;88f75481-0a52-11de-6fe3-fe72f7be4be1;8c1f9691-0a52-11de-6fe3-fe72f7be4be1;8a407651-0a52-11de-6fe3-fe72f7be4be1;92f0e371-0a52-11de-6fe3-fe72f7be4be1;90ceedd1-0a52-11de-6fe3-fe72f7be4be1;9801e2b1-0a52-11de-6fe3-fe72f7be4be1;9c64c7a1-0a52-11de-6fe3-fe72f7be4be1;8e0ac4c1-0a52-11de-6fe3-fe72f7be4be1;895b9491-0a52-11de-6fe3-fe72f7be4be1;89c6d981-0a52-11de-6fe3-fe72f7be4be1;96af4b01-0a52-11de-6fe3-fe72f7be4be1;8ee87a91-0a52-11de-6fe3-fe72f7be4be1;91d78431-0a52-11de-6fe3-fe72f7be4be1;99678d31-0a52-11de-6fe3-fe72f7be4be1;887441d1-0a52-11de-6fe3-fe72f7be4be1;8b71a351-0a52-11de-6fe3-fe72f7be4be1;9568c141-0a52-11de-6fe3-fe72f7be4be1;9e7d4761-0a52-11de-6fe3-fe72f7be4be1;80ca2121-0a52-11de-6fe3-fe72f7be4be1;83da6e61-0a52-11de-6fe3-fe72f7be4be1;820beff1-0a52-11de-6fe3-fe72f7be4be1;829b1091-0a52-11de-6fe3-fe72f7be4be1;8333a711-0a52-11de-6fe3-fe72f7be4be1;81297f31-0a52-11de-6fe3-fe72f7be4be1;807dd5e1-0a52-11de-6fe3-fe72f7be4be1;8194c421-0a52-11de-6fe3-fe72f7be4be1;80422c71-0a52-11de-6fe3-fe72f7be4be1;8709dc61-0a52-11de-6fe3-fe72f7be4be1;86c24c11-0a52-11de-6fe3-fe72f7be4be1;87b314b1-0a52-11de-6fe3-fe72f7be4be1;8753b6a1-0a52-11de-6fe3-fe72f7be4be1;851c6441-0a52-11de-6fe3-fe72f7be4be1;85559cb1-0a52-11de-6fe3-fe72f7be4be1;85fc6401-0a52-11de-6fe3-fe72f7be4be1;85a1e7f1-0a52-11de-6fe3-fe72f7be4be1;79297651-0a52-11de-6fe3-fe72f7be4be1;73f95651-0a52-11de-6fe3-fe72f7be4be1;5db98951-0a52-11de-6fe3-fe72f7be4be1;649b7801-0a52-11de-6fe3-fe72f7be4be1;6145e461-0a52-11de-6fe3-fe72f7be4be1;6b338c71-0a52-11de-6fe3-fe72f7be4be1;6699b901-0a52-11de-6fe3-fe72f7be4be1;5df077d1-0a52-11de-6fe3-fe72f7be4be1;725f5561-0a52-11de-6fe3-fe72f7be4be1;6de95da1-0a52-11de-6fe3-fe72f7be4be1;639c5781-0a52-11de-6fe3-fe72f7be4be1;67a70a51-0a52-11de-6fe3-fe72f7be4be1;5efdc921-0a52-11de-6fe3-fe72f7be4be1;609f1d11-0a52-11de-6fe3-fe72f7be4be1;6c83b321-0a52-11de-6fe3-fe72f7be4be1;7b0b0791-0a52-11de-6fe3-fe72f7be4be1;658c40a1-0a52-11de-6fe3-fe72f7be4be1;5e3a5211-0a52-11de-6fe3-fe72f7be4be1;62c83ea1-0a52-11de-6fe3-fe72f7be4be1;5e94f531-0a52-11de-6fe3-fe72f7be4be1;68c79581-0a52-11de-6fe3-fe72f7be4be1;70d13b51-0a52-11de-6fe3-fe72f7be4be1;60041591-0a52-11de-6fe3-fe72f7be4be1;7cef09d1-0a52-11de-6fe3-fe72f7be4be1;5f74f4f1-0a52-11de-6fe3-fe72f7be4be1;775d41d1-0a52-11de-6fe3-fe72f7be4be1;62025691-0a52-11de-6fe3-fe72f7be4be1;6f53c311-0a52-11de-6fe3-fe72f7be4be1;69f67891-0a52-11de-6fe3-fe72f7be4be1;75a1d631-0a52-11de-6fe3-fe72f7be4be1;5c079391-0a52-11de-6fe3-fe72f7be4be1;5bbdb951-0a52-11de-6fe3-fe72f7be4be1;5ccd7ba1-0a52-11de-6fe3-fe72f7be4be1;5c5fc5b1-0a52-11de-6fe3-fe72f7be4be1;5b7fc5f1-0a52-11de-6fe3-fe72f7be4be1;4c8f7681-0a52-11de-6fe3-fe72f7be4be1;4e868b91-0a52-11de-6fe3-fe72f7be4be1;4ae71db1-0a52-11de-6fe3-fe72f7be4be1;48df66d1-0a52-11de-6fe3-fe72f7be4be1;49aec4c1-0a52-11de-6fe3-fe72f7be4be1;4d4977b1-0a52-11de-6fe3-fe72f7be4be1;4a1798b1-0a52-11de-6fe3-fe72f7be4be1;495692a1-0a52-11de-6fe3-fe72f7be4be1;4bdf1241-0a52-11de-6fe3-fe72f7be4be1;4ffa66e1-0a52-11de-6fe3-fe72f7be4be1;4f7c0f21-0a52-11de-6fe3-fe72f7be4be1;48bb8b21-0a52-11de-6fe3-fe72f7be4be1;4b846f21-0a52-11de-6fe3-fe72f7be4be1;4a93a681-0a52-11de-6fe3-fe72f7be4be1;491fa421-0a52-11de-6fe3-fe72f7be4be1;4e11a9b1-0a52-11de-6fe3-fe72f7be4be1;7f11f461-0a53-11de-6fe3-fe72f7be4be1;7f4b2cd1-0a53-11de-6fe3-fe72f7be4be1;802d9d91-0a53-11de-6fe3-fe72f7be4be1;85187731-0a53-11de-6fe3-fe72f7be4be1;c7756853-0a52-11de-6fe3-fe72f7be4be1;d89ac031-0a52-11de-6fe3-fe72f7be4be1;d6f96c43-0a52-11de-6fe3-fe72f7be4be1;d1e14111-0a52-11de-6fe3-fe72f7be4be1;d8105a81-0a52-11de-6fe3-fe72f7be4be1;cc94bf73-0a52-11de-6fe3-fe72f7be4be1;c79b8df1-0a52-11de-6fe3-fe72f7be4be1;c819e5b3-0a52-11de-6fe3-fe72f7be4be1;d247cb11-0a52-11de-6fe3-fe72f7be4be1;c95e2583-0a52-11de-6fe3-fe72f7be4be1;e54eb433-0a52-11de-6fe3-fe72f7be4be1;ec2be7f1-0a52-11de-6fe3-fe72f7be4be1;ea1a9421-0a52-11de-6fe3-fe72f7be4be1;e580c0b1-0a52-11de-6fe3-fe72f7be4be1;e739e261-0a52-11de-6fe3-fe72f7be4be1;77a609e5-0811-11de-6fe3-fe72f7be4be1;edd6b1c1-0a52-11de-6fe3-fe72f7be4be1;ee0b2f41-0a52-11de-6fe3-fe72f7be4be1;59aed683-0a52-11de-6fe3-fe72f7be4be1;59c43341-0a52-11de-6fe3-fe72f7be4be1;5a36a423-0a52-11de-6fe3-fe72f7be4be1;5aa91503-0a52-11de-6fe3-fe72f7be4be1;1ba5a491-0a52-11de-6fe3-fe72f7be4be1;1bacd083-0a52-11de-6fe3-fe72f7be4be1;327bbe21-0a52-11de-6fe3-fe72f7be4be1;51248f01-0a52-11de-6fe3-fe72f7be4be1;5137a1d1-0a52-11de-6fe3-fe72f7be4be1;33f6c563-0a52-11de-6fe3-fe72f7be4be1;46330b81-0a52-11de-6fe3-fe72f7be4be1;46999581-0a52-11de-6fe3-fe72f7be4be1;4268bd11-0a52-11de-6fe3-fe72f7be4be1;3d850f61-0a52-11de-6fe3-fe72f7be4be1;34006253-0a52-11de-6fe3-fe72f7be4be1;5b441c83-0a52-11de-6fe3-fe72f7be4be1;a7b5a931-0a52-11de-6fe3-fe72f7be4be1;a2e02c51-0a52-11de-6fe3-fe72f7be4be1;884e1c31-0a52-11de-6fe3-fe72f7be4be1;80174be3-0a52-11de-6fe3-fe72f7be4be1;869c2671-0a52-11de-6fe3-fe72f7be4be1;84f88891-0a52-11de-6fe3-fe72f7be4be1;5d95d4b3-0a52-11de-6fe3-fe72f7be4be1;5b5e5b43-0a52-11de-6fe3-fe72f7be4be1;48897ea3-0a52-11de-6fe3-fe72f7be4be1;489c9171-0a52-11de-6fe3-fe72f7be4be1;1b760911-0a52-11de-6fe3-fe72f7be4be1;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;1b5a4a4a-080e-11de-6fe3-fe72f7be4be1;1b5a4a4e-080e-11de-6fe3-fe72f7be4be1;1b5a4a50-080e-11de-6fe3-fe72f7be4be1;1b5a4a4c-080e-11de-6fe3-fe72f7be4be1;38059a21-04b0-11de-6fe3-fe72f7be4be1;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;40a4eb40-f13a-11dd-4d05-5048494f4e43;40a4eb41-f13a-11dd-4d05-5048494f4e43;1837e460-f140-11dd-4d05-5048494f4e43;1837e461-f140-11dd-4d05-5048494f4e43;fbaadeb5-04b3-11de-6fe3-fe72f7be4be1;1837e462-f140-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40cd-e711-11dd-4d05-5048494f4e43;4a4d40cc-e711-11dd-4d05-5048494f4e43;4a4d40ce-e711-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;34423d00-eca8-11dd-4d05-5048494f4e43;76b21d68-a372-4d83-9c5f-927cb14a9166;b031acec-a14d-46db-82e2-b30983e49400;726fd210-c768-11dd-6d8b-bf08ef135bea;e0deef48-aeb9-406e-a253-61addeb62734;1016a034-4b50-41e6-a93e-500a417632bd;726fd211-c768-11dd-6d8b-bf08ef135bea;62593d17-617d-47c4-a98e-67777052639e;385d8c2d-1d54-46a7-930f-9b90b28c9383;44e27748-ff5b-45d8-abb1-86239a6be68d;d6f37538-a3e5-11dc-067d-000ae482fe4d;05159f3c-16f1-44bc-962f-e0963cb74614;910a6117-2508-4eed-85f2-81b46f2ae967;9eb4f8d0-2d47-412d-b9c4-29f4b31ab576;09b75ad5-be6e-48e1-96b1-6d479646cbd6;c425bf17-063d-4313-b5f7-f4a25a5ab6ae;bdd36de0-c828-11dd-6d8b-bf08ef135bea;607b968c-a7ac-4f30-a743-f2678e69a995;db814fa9-9056-46fc-9f2b-155abfe328db;cc89c749-dcf6-4940-95bc-4618b908f068;726fd224-c768-11dd-6d8b-bf08ef135bea;79b3bf70-0695-45b8-9b01-cc5f177c099c;19c12714-da03-4010-a6f5-01580086095b;97499dfd-49be-4348-8322-dde781cfa28a;3ecad984-d87d-4728-b940-2e2de25b5bdd;118b9b96-9c28-4fa8-9194-f620d0d9f4cd;29dedb7a-6160-40ac-afa6-ee9f4071f252;d6f37530-a3e5-11dc-067d-000ae482fe4d;726fd228-c768-11dd-6d8b-bf08ef135bea;726fd229-c768-11dd-6d8b-bf08ef135bea;5122c41d-08ca-481e-bf88-2d92927a958f;19e22f1f-5415-4a37-9c3a-38b5ed7da413;3a839ab5-e47f-479b-a11a-f3a05c579d17;865f65a3-15b5-4829-b9a8-1f01cbb1bd7a;40052cf2-ebff-4e7b-8645-8ebff20f484e;9f8ef3ef-e61b-4a34-95bc-0b4268528364;7c6a8c73-c128-44b7-a79c-95ef52772568;69bbeaef-1fc5-4b20-9fe2-eb8822e83764;d6f3753f-a3e5-11dc-067d-000ae482fe4d;d6f37540-a3e5-11dc-067d-000ae482fe4d;8483dab8-7888-4a9f-8453-3b801b8c1438;");                
                if (exportRelevantObjectGUIDsPropertyValue != null && exportRelevantObjectGUIDsPropertyValue.length() > 0) {
                    var input = exportRelevantObjectGUIDsPropertyValue.split(";");
                    for (var i = 0; i < input.length; i++) {
                        g_oSelObjDefs.push(ArisData.getActiveDatabase().FindGUID(input[i]));
                    }
                }
            }
            
            //IV. ----
            //Abort if this error condition is met: there are risks that are generics and specilaizations at the same time
            if (g_error_risk_genericSpecific.size() > 0) {
                var aErroneousRisksArray = convertHashSetToJSArray(g_error_risk_genericSpecific);
                var sErrorMsg = getString("TEXT_4");
                for (var i=0; i<aErroneousRisksArray.length; i++) {
                    sErrorMsg += aErroneousRisksArray[i].Name(g_nLoc);
                    if (i<aErroneousRisksArray.length - 1) {sErrorMsg += "\n";}
                }
                
                if (g_bRemoteCalled) {
                    Context.setProperty(g_errorMessageProperty, sErrorMsg);
                } else {
                    Dialogs.MsgBox(sErrorMsg, Constants.MSGBOX_BTN_OK | Constants.MSGBOX_ICON_ERROR, getString("TEXT_3"));
                }
                return;
            }
            //Abort if this error condition is met: there are hierarchy elements that both seem to belong to orgunit and tester hierarchy
            var hierarchyIntersection = new java.util.HashSet(g_orgUnitSet);
            hierarchyIntersection.retainAll(g_testerSet);
            if (hierarchyIntersection.size() > 0) {
                var aErroneousHierarchyElements = convertHashSetToJSArray(hierarchyIntersection);
                var sErrorMsg = getString("TEXT_5");
                for (var i=0; i<aErroneousHierarchyElements.length; i++) {
                    sErrorMsg = sErrorMsg + "- " + aErroneousHierarchyElements[i].Name(g_nLoc);
                    if (i<aErroneousHierarchyElements.length - 1) {sErrorMsg += "\n";}
                }
                
                if (g_bRemoteCalled) {
                    Context.setProperty(g_errorMessageProperty, sErrorMsg);
                } else {
                    Dialogs.MsgBox(sErrorMsg, Constants.MSGBOX_BTN_OK | Constants.MSGBOX_ICON_ERROR, getString("TEXT_3"));
                }
                return;        
            }
          
            //V. ----
            //generate XML for the found objects from the mapping file
            var xmlOutput = Context.createXMLOutputObject(Context.getSelectedFile(), "import");
            generateXMLExport(xmlOutput, g_oSelObjDefs);
			
			//write the generated XML
			xmlOutput.WriteReport();
                   
        }

    } catch(e) {       
        var sErrorMsg = e.toString();
        sErrorMsg += "\nfileName: " + e.fileName;
        sErrorMsg += "\nlineNumber: " + e.lineNumber;
       
        if (g_bRemoteCalled) {
            Context.setProperty(g_errorMessageProperty, sErrorMsg);
        } else {
            Dialogs.MsgBox(sErrorMsg, Constants.MSGBOX_BTN_OK | Constants.MSGBOX_ICON_ERROR, getString("TEXT_3"));
        }
    }

}


/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
          ~~~~~~~ Functionality for XML creation from mapping file ~~~~~~~
  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
/*---------------------------------------------------------------------------------------
    Determine the origin for this export
---------------------------------------------------------------------------------------*/
function determineOrigin() {
    var oOriginDatabase = null;
    if (ArisData.getSelectedDatabases().length != 0) {
        oOriginDatabase = ArisData.getSelectedDatabases()[0];
    } else if (ArisData.getSelectedGroups().length != 0) {
        oOriginDatabase = ArisData.getSelectedGroups()[0].Database();
    } else if (ArisData.getSelectedModels().length != 0) {
        oOriginDatabase = ArisData.getSelectedModels()[0].Database();
    } else {
        oOriginDatabase = ArisData.getSelectedObjDefs()[0].Database();
    }
    return oOriginDatabase.Name(g_nLoc);
}

/*---------------------------------------------------------------------------------------
    Main XML export function
---------------------------------------------------------------------------------------*/
function generateXMLExport(xmlOutput) {

    //--- determine origin
    var sOrigin = determineOrigin();
    
    //--- determine scope
    var sScope = "PARTIAL";
    if (g_bImportModeEntire == "true") {
        sScope = "ENTIRE";
    }
    
	//--- XML format
    var xmlOutput_Root = xmlOutput.getRootElement();
    
    var xmlConfig_Root = g_xmlConfigReader.getRootElement();
 
	//create HashMap for accessing <mappingObject>s by obj_Type and aris_typenum
    xmlConfig_Mapping = xmlConfig_Root.getChild("aris2aamMapping");
	hm_MappingObjects_Arcm2Aris = createHashMap_MappingObjects(xmlConfig_Mapping);
    
    //write the overall meta data for logging on ARCM side
    generateOverallMetaData(xmlOutput, xmlOutput_Root, xmlConfig_Mapping, sOrigin, sScope);
    
	// --- <mappingObject> -> <Objects>, <Object> ---
    //iterate over all defined <mappingObject>s for creating XML content
    var iterObj = xmlConfig_Mapping.getChildren("mappingObject").iterator(); 
    while (iterObj.hasNext()) {
		var xmlConfig_ObjectTypeMapping = iterObj.next();
        generateXMLExport_MappingObject(xmlOutput, xmlOutput_Root, xmlConfig_ObjectTypeMapping, hm_MappingObjects_Arcm2Aris)
    }
    
    //set quantity after creating the XML content
    xmlOutput_Root.setAttribute("quantity", g_objDefXMLOutputObjectMap.size());
}

/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
          ~~~~~~~ Functionality for export XML - meta data creation ~~~~~~~
  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
/*---------------------------------------------------------------------------------------
    XML export main function for overall meta data
---------------------------------------------------------------------------------------*/
function generateOverallMetaData(xmlOutput, xmlOutput_RootNode, xmlConfig_Mapping, sOrigin, sScope) {
    
    //skip in debugging mode
    if (g_bDebug) {return;}
    
    //<meta>
    var xmlOutput_Meta = xmlOutput.addElement(xmlOutput_RootNode, "meta");
    //<inf>
    var xmlOutput_Meta_Inf = xmlOutput.addElement(xmlOutput_Meta, "inf");
    xmlOutput_Meta_Inf.setAttribute("id", "origin");
    xmlOutput_Meta_Inf.setText(sOrigin);
    
    //<inf>
    xmlOutput_Meta_Inf = xmlOutput.addElement(xmlOutput_Meta, "inf");
    xmlOutput_Meta_Inf.setAttribute("id", "scope");
    xmlOutput_Meta_Inf.setText(sScope);
    
    //<mapping>
    var xmlOutput_Meta_Mapping = xmlOutput.addElement(xmlOutput_Meta, "mapping");
    
    var iterObj = xmlConfig_Mapping.getChildren("mappingObject").iterator();
    while (iterObj.hasNext()) {
        var xmlConfig_MappingObject = iterObj.next();
        var sObjType = xmlConfig_MappingObject.getAttribute("obj_type").getValue();
        
        //meta data cache
        var oAlreadyCreatedAttributesSet = hmMeta_ObjTypeID2attributeIDs.get(sObjType);
        if (oAlreadyCreatedAttributesSet == null) {
            oAlreadyCreatedAttributesSet = new java.util.HashSet();
            hmMeta_ObjTypeID2attributeIDs.put(sObjType, oAlreadyCreatedAttributesSet);
        }
        
        //<object> - either new or from cache
        var xmlOutput_Meta_Mapping_Object = hmMeta_ObjTypeID2outputNode.get(sObjType);
        if (xmlOutput_Meta_Mapping_Object == null) {
            xmlOutput_Meta_Mapping_Object = xmlOutput.addElement(xmlOutput_Meta_Mapping, "object");
            xmlOutput_Meta_Mapping_Object.setAttribute("id", sObjType);
            hmMeta_ObjTypeID2outputNode.put(sObjType, xmlOutput_Meta_Mapping_Object);
        }
         
        var iterAttr = xmlConfig_MappingObject.getChildren("mappingAttr").iterator();
        while (iterAttr.hasNext()) {
            var xmlConfig_MappingAttr = iterAttr.next();
            var sName = xmlConfig_MappingAttr.getAttribute("name").getValue();            
            //only process those attributes for which there were no meta data created already 
            if (!oAlreadyCreatedAttributesSet.add(sName)) {
                continue;
            }
            var sArisTypeNum = xmlConfig_MappingAttr.getAttribute("aris_typenum").getValue();
            var sLocalizedAttrName = "";
            //handle multi enum attributes by resolving their enum mapping attributes
            if (sArisTypeNum == "ISMULTIPLE") {
                var sEnumType = xmlConfig_MappingAttr.getAttribute("type").getValue();
                var xmlConfig_MappingEnumElement = getXmlConfigMappingEnumElement(sEnumType);
                if (xmlConfig_MappingEnumElement != null) {
                    var xmlConfig_EnumItems = xmlConfig_MappingEnumElement.getChildren("enumItem");
                    var iterItems = xmlConfig_EnumItems.iterator();
                    while (iterItems.hasNext()) {
                        var xmlConfig_EnumItem = iterItems.next();
                        var sLocalizedEnumPartAttrTypeNum = getAttributeTypeNum(xmlConfig_EnumItem.getAttribute("aris_typenum").getValue());
                        if (sLocalizedAttrName != "") {sLocalizedAttrName += " / ";}
                        sLocalizedAttrName += localizeMetaData(sLocalizedEnumPartAttrTypeNum, "ATTRIBUTE");
                    }
                }    
            }
            //handle all other value attributes
            else {
                sLocalizedAttrName = localizeMetaData(sArisTypeNum, "ATTRIBUTE");
            }
            
            //only create <attr> and <inf>  if attribute type could be resolved and localized
            if (sLocalizedAttrName != sArisTypeNum) {
                //<attr>
                var xmlOutput_Meta_Mapping_Object_Attr = xmlOutput.addElement(xmlOutput_Meta_Mapping_Object, "attr");
                xmlOutput_Meta_Mapping_Object_Attr.setAttribute("id", sName);
                //<inf>
                var xmlOutput_Meta_Mapping_Object_Attr_Inf = xmlOutput.addElement(xmlOutput_Meta_Mapping_Object_Attr, "inf");
                xmlOutput_Meta_Mapping_Object_Attr_Inf.setAttribute("id", "aris_typenum");
                xmlOutput_Meta_Mapping_Object_Attr_Inf.setText(sLocalizedAttrName);
            }
        }
        
        var iterLink = xmlConfig_MappingObject.getChildren("mappingLink").iterator();
        while (iterLink.hasNext()) {
            var xmlConfig_MappingLink = iterLink.next();
            var sID = xmlConfig_MappingLink.getAttribute("id").getValue();
            //only process those attributes for which there were no meta data created already 
            if (!oAlreadyCreatedAttributesSet.add(sID)) {
                continue;
            }
            var sArisTypeNum = xmlConfig_MappingLink.getAttribute("aris_typenum").getValue();
            var sDirection = "";
            if (xmlConfig_MappingLink.getAttribute("direction") != null) {sDirection = xmlConfig_MappingLink.getAttribute("direction").getValue();}
            var sTargetObjType = "";
            if (xmlConfig_MappingLink.getAttribute("aris_target_objtype") != null) {sTargetObjType = xmlConfig_MappingLink.getAttribute("aris_target_objtype").getValue();}
            var sTargetSymbolNum = "";
            if (xmlConfig_MappingLink.getAttribute("aris_target_symbolnum") != null) {sTargetSymbolNum = xmlConfig_MappingLink.getAttribute("aris_target_symbolnum").getValue();}

            var sLocalizedCxnName = localizeMetaData(sArisTypeNum, "CONNECTION");
            var sLocalizedTargetObjType = localizeMetaData(sTargetObjType, "OBJDEF");
            var sLocalizedTargetSymbolNum = localizeMetaData(sTargetSymbolNum, "SYMBOL");
            //skip if there would be nothing to add to output
            if ((sLocalizedCxnName == sArisTypeNum) && (sDirection == "") && (sTargetObjType == "" || sLocalizedTargetObjType == sTargetObjType)) {
                continue;
            }
            
            //<attr>
            var xmlOutput_Meta_Mapping_Object_Attr = xmlOutput.addElement(xmlOutput_Meta_Mapping_Object, "attr");
            xmlOutput_Meta_Mapping_Object_Attr.setAttribute("id", sID);
            //<inf> 
            if (sLocalizedCxnName != sArisTypeNum) {
                var xmlOutput_Meta_Mapping_Object_Attr_Inf = xmlOutput.addElement(xmlOutput_Meta_Mapping_Object_Attr, "inf");
                xmlOutput_Meta_Mapping_Object_Attr_Inf.setAttribute("id", "aris_typenum");
                xmlOutput_Meta_Mapping_Object_Attr_Inf.setText(sLocalizedCxnName);
            }
            //<inf>
            if (sDirection != "") {
                xmlOutput_Meta_Mapping_Object_Attr_Inf = xmlOutput.addElement(xmlOutput_Meta_Mapping_Object_Attr, "inf");
                xmlOutput_Meta_Mapping_Object_Attr_Inf.setAttribute("id", "direction");
                xmlOutput_Meta_Mapping_Object_Attr_Inf.setText(sDirection);
            }
            
            //<inf> - only if ObjDef type can be resolved and localized      
            if (sTargetObjType != "" && sLocalizedTargetObjType != sTargetObjType) {
                xmlOutput_Meta_Mapping_Object_Attr_Inf = xmlOutput.addElement(xmlOutput_Meta_Mapping_Object_Attr, "inf");
                xmlOutput_Meta_Mapping_Object_Attr_Inf.setAttribute("id", "aris_target_objtype");
                xmlOutput_Meta_Mapping_Object_Attr_Inf.setText(sLocalizedTargetObjType);
            }
            
            //<inf> - only if symbol type can be resolved and localized    
            if (sTargetSymbolNum != "" && sLocalizedTargetSymbolNum != sTargetSymbolNum) {
                xmlOutput_Meta_Mapping_Object_Attr_Inf = xmlOutput.addElement(xmlOutput_Meta_Mapping_Object_Attr, "inf");
                xmlOutput_Meta_Mapping_Object_Attr_Inf.setAttribute("id", "aris_target_symbolnum");
                xmlOutput_Meta_Mapping_Object_Attr_Inf.setText(sLocalizedTargetSymbolNum);
            }
        }
    }
}

/*---------------------------------------------------------------------------------------
    XML export main function for overall meta data - localization
---------------------------------------------------------------------------------------*/
function localizeMetaData(sInput, sMode) {
    var sResult = sInput;
    //attribute name by attrTypeNum
    if (sMode == "ATTRIBUTE") {
        var iTypeNum = obtainTypeNumFromFilter(sInput, "UserDefinedAttributeTypeNum");
        if (iTypeNum != -1) {
            sResult = ArisData.ActiveFilter().getItemTypeName(Constants.CID_ATTRDEF, iTypeNum);
        }
    }
    //attribute name by cxnTypeNum
    if (sMode == "CONNECTION") {
        var iTypeNum = obtainTypeNumFromFilter(sInput, null);
        if (iTypeNum != -1) {
            sResult = ArisData.ActiveFilter().getItemTypeName(Constants.CID_CXNDEF, iTypeNum);
        }
    }
    //target ObjDef type
    if (sMode == "OBJDEF") {
        var iTypeNum = obtainTypeNumFromFilter(sInput, null);
        if (iTypeNum != -1) {
            sResult = ArisData.ActiveFilter().getItemTypeName(Constants.CID_OBJDEF, iTypeNum);
        }
    }
    //target Symbol type
    if (sMode == "SYMBOL") {
        var iTypeNum = obtainTypeNumFromFilter(sInput, "UserDefinedSymbolTypeNum");
        if (iTypeNum != -1) {
            sResult = ArisData.ActiveFilter().SymbolName(iTypeNum);
        }
    }
    
    return sResult;
}

/*---------------------------------------------------------------------------------------
    XML export function for object specific meta data 
---------------------------------------------------------------------------------------*/
function writeXmlElement_ObjectMetaData(xmlOutput, xmlOutput_ObjectNode, oObjDef) {
    var oGroup = oObjDef.Group();
    var sPath = oGroup.Path(g_nLoc, true, true);
    var sName = oObjDef.Name(g_nLoc);
    var xmlOutput_Meta = xmlOutput.addElement(xmlOutput_ObjectNode, "meta");
    
    var xmlOutput_Meta_Inf = xmlOutput.addElement(xmlOutput_Meta, "inf");
    xmlOutput_Meta_Inf.setAttribute("id", "group");
    xmlOutput_Meta_Inf.setText(sPath);
    
    xmlOutput_Meta_Inf = xmlOutput.addElement(xmlOutput_Meta, "inf");
    xmlOutput_Meta_Inf.setAttribute("id", "name");
    xmlOutput_Meta_Inf.setText(sName);
}

/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
          ~~~~~~~ Functionality for export XML - meta data creation ~~~~~~~
  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
/*-----------------------------------------------------------------------------------------------------------
    XML export function for all ObjDefs targeted by a specific <mappingObject> element from mapping file
-----------------------------------------------------------------------------------------------------------*/
function generateXMLExport_MappingObject(xmlOutput, xmlOutput_ImportNode, xmlConfig_ObjectTypeMapping, hm_MappingObjects_Arcm2Aris) {
        
	//read mapping object attributes
	if (xmlConfig_ObjectTypeMapping.getAttribute("obj_type") == null) {return;}
	var arcmType = xmlConfig_ObjectTypeMapping.getAttribute("obj_type").getValue();
	
	var sTypeNum = "";
	if (xmlConfig_ObjectTypeMapping.getAttribute("aris_typenum") != null) {sTypeNum = xmlConfig_ObjectTypeMapping.getAttribute("aris_typenum").getValue();}       
	var sSymbolNum = "";
	if (xmlConfig_ObjectTypeMapping.getAttribute("aris_symbolnum") != null) {sSymbolNum = xmlConfig_ObjectTypeMapping.getAttribute("aris_symbolnum").getValue();}        
	var sInfo = "";
    if (xmlConfig_ObjectTypeMapping.getAttribute("info") != null) {sInfo = xmlConfig_ObjectTypeMapping.getAttribute("info").getValue();}
    var sMappingApproach = "";
    if (xmlConfig_ObjectTypeMapping.getAttribute("approach") != null) {sMappingApproach = xmlConfig_ObjectTypeMapping.getAttribute("approach").getValue();}
    
	//skip if there is no objTypeNum
	if (sTypeNum == null || sTypeNum.length() == 0) {
		return;
	}
    
    //skip if mapping object belongs to a different approach
    if (sMappingApproach != "" && sMappingApproach != g_approach) {
        return;
    }
	
	//if objTypeNum is defined: get the qualified objects for the given type
	var aObjDefs = getObjectsByType(g_oSelObjDefs, sTypeNum, sSymbolNum, arcmType);
	//if there are no objects for this object type then check the next mapping entry
	if (aObjDefs.length == 0) {
		return;
	}        
	
	//skip policy review task tag, it is handled together with policy
	if (arcmType == "POLICYREVIEWTASK") {
		return;
	}
	
	//skip standard section and question mapping, they are handled together with questionnaire templates
    if (arcmType == "SECTION" || arcmType == "QUESTION") {
		return;
	}
    
	//skip audit step templates, they are handled together with audit templates
	if (arcmType == "AUDITSTEPTEMPLATE") {    
		return;
	}
    
    //sort hierarchies to leaf-first - this is needed for the export file consistency
    if (arcmType == "HIERARCHY") {
        aObjDefs = sortHierarchiesChildrenFirst(aObjDefs);
    }
	
	//create XML for all qualified objects
	if (aObjDefs.length > 0) {
		for (var i = 0 ; i < aObjDefs.length ; i++ ) {
			var oObjDef = aObjDefs[i];
			generateXMLExport_ObjDef(xmlOutput, xmlOutput_ImportNode, xmlConfig_ObjectTypeMapping, hm_MappingObjects_Arcm2Aris, oObjDef, arcmType, sInfo)
		}
	}	
}

/*----------------------------------------------------------
    XML export function for single element <Object>
----------------------------------------------------------*/
function generateXMLExport_ObjDef(xmlOutput, xmlOutput_ImportNode, xmlConfig_ObjectTypeMapping, hm_MappingObjects_Arcm2Aris, oObjDef, arcmType, sInfo) {

	// SPECIAL CASE: Documents
	if (arcmType == "DOCUMENT") {
		createXmlElement_Document(xmlConfig_ObjectTypeMapping, xmlOutput, xmlOutput_ImportNode, oObjDef);
	}
	
	// SPECIAL CASE: Policy review task for and policy definition
	else if (arcmType == "POLICYDEFINITION") {
        createXmlElement_PolicyDefinition_PolicyReviewTask(xmlConfig_ObjectTypeMapping, xmlOutput, xmlOutput_ImportNode, oObjDef);
        return;
	}
	
    // SPECIAL CASE: for each questionnaire template export separate copies of the sections (in recursion) 
	else if (arcmType == "QUESTIONNAIRE_TEMPLATE") {
		createXmlElement_QuestionnaireTemplate(xmlOutput, xmlOutput_ImportNode, hm_MappingObjects_Arcm2Aris, oObjDef, arcmType);
	}
	
	// SPECIAL CASE: for each audit template export separate copies of the step templates (in recursion) 
	else if (arcmType == "AUDITTEMPLATE") {
		createXmlElement_AuditTemplate(xmlOutput, xmlOutput_ImportNode, hm_MappingObjects_Arcm2Aris, oObjDef, arcmType);
	}
	
	// DEFAULT:
	else {
		// write current object
		var aObjList2CheckCycles = new Array();  // List with all children for checking cycles
		aObjList2CheckCycles.push(oObjDef);
		createXmlElement_StandardObject(xmlOutput, xmlOutput_ImportNode, xmlConfig_ObjectTypeMapping, oObjDef, aObjList2CheckCycles, arcmType, "");
	}

}


/*---------------------------------------------------------------------------------------
    Build HashMap for accessing <mappingObject>s by obj_Type, aris_typenum and info
	Example: USER|OT_PERS
    Exapmle: SECTION|OT_SURVEY_SECTION|dependentVisibility
---------------------------------------------------------------------------------------*/
function createHashMap_MappingObjects(xmlConfig_Mapping) {
    var hm_MappingObjects_Arcm2Aris = new java.util.HashMap();          
    var xmlConfig_MappingObjects = xmlConfig_Mapping.getChildren("mappingObject");
    var iterObj = xmlConfig_MappingObjects.iterator();
    while (iterObj.hasNext()) {
        var xmlConfig_ObjectTypeMapping = iterObj.next();
        if (xmlConfig_ObjectTypeMapping.getAttribute("obj_type").getValue() != null) {
            var key = xmlConfig_ObjectTypeMapping.getAttribute("obj_type").getValue();
            if (xmlConfig_ObjectTypeMapping.getAttribute("aris_typenum").getValue() != null) {
                key = key + "|" + xmlConfig_ObjectTypeMapping.getAttribute("aris_typenum").getValue();
            }
            if (xmlConfig_ObjectTypeMapping.getAttribute("info") != null && xmlConfig_ObjectTypeMapping.getAttribute("info").getValue() != null) {
                key = key + "|" + xmlConfig_ObjectTypeMapping.getAttribute("info").getValue();
            }
            hm_MappingObjects_Arcm2Aris.put(key, xmlConfig_ObjectTypeMapping);
        }    
    }
	return hm_MappingObjects_Arcm2Aris;
}



/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
          ~~~~~~~ Functionality for creation of XML elements ~~~~~~~
  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
  
/*---------------------------------------------------------------------------------------
    Output of documents:
    - No check of multiple output
    - Attributes: only title and description are output if both are maintained
    - Links: no links considered
---------------------------------------------------------------------------------------*/
function createXmlElement_Document(xmlConfig_ObjectTypeMapping, xmlOutput, xmlOutput_Objects, oObjDef) {
    
    // --- mappingAttr -> Attr ---
    var aAttrElements = createXmlElement_Common_MappingAttributes(xmlConfig_ObjectTypeMapping, oObjDef, false, false, true);
        
	// "type" is always included in export so the attribute count must be 2 or more
    if (aAttrElements.length >= 2) {
          
        //update GUID mapping caches
        var sListAttributeID = xmlConfig_ObjectTypeMapping.getAttribute("parent_link_id").getValue();
        var sDocName = ""
        var sDocLink = "";
        for (var i=0; i<aAttrElements.length; i++) {
            if (aAttrElements[i].name == "name") {sDocName = aAttrElements[i].value;}
            if (aAttrElements[i].name == "link") {sDocLink = aAttrElements[i].value;}
        }
        updateDocumentCache(oObjDef.GUID(), sListAttributeID, sDocName, sDocLink);
 
        //determine doc GUID in ARCM
        var docGUID = getDocumentGUIDFromGUIDMappingCache(sDocName, sDocLink, oObjDef.GUID());
        m_aDocIdList.push(new documentID(oObjDef, docGUID));
     
        //-----------------------------
        // out <object>
        var xmlOutput_Object = writeXmlElement_Object(xmlOutput, xmlOutput_Objects, oObjDef, docGUID, "DOCUMENT");
        
        //-----------------------------
        // out <meta>
        writeXmlElement_ObjectMetaData(xmlOutput, xmlOutput_Object, oObjDef);
        
        //-----------------------------
        // out <attr>s
        writeXmlElement_Attributes(xmlOutput, xmlOutput_Object, aAttrElements);
    }
}

/*---------------------------------------------------------------------------------------------------
	Output for a policy review task and a policy definition created from a given policy definition.
----------------------------------------------------------------------------------------------------*/
function createXmlElement_PolicyDefinition_PolicyReviewTask(xmlConfig_ObjectTypeMapping_pd, xmlOutput, xmlOutput_ImportNode, p_oPolicyDefinition) {

    //===========================================
    //=== first export the policy review task ===
    
    var xmlConfig_ObjectTypeMapping_prt = hm_MappingObjects_Arcm2Aris.get("POLICYREVIEWTASK|OT_POLICY");
    
	// --- mappingAttr -> Attr ---
    var aAttrElements = createXmlElement_Common_MappingAttributes(xmlConfig_ObjectTypeMapping_prt, p_oPolicyDefinition, true, false, false);

    // --- mappingLink -> Link ---
    // hard coded since policy review tasks do not exist in ARIS and the linked objects must be taken from the policy definition
    var aLinkElements = new Array();
    
    //link to policy definition's owner
    aLinkElements = aLinkElements.concat( createXmlElement_LinksByHashMap(p_oPolicyDefinition, "owner_group", g_policyDefinition2policyOwnerHashMap) );
    
    //update GUID mapping caches
    updatePolicyReviewTaskCache(p_oPolicyDefinition.GUID());

    //determine policy review task GUID in ARCM
    var task_guid = getPolicyReviewTaskGUIDFromGUIDMappingCache(p_oPolicyDefinition.GUID());
      
    //----------------------------------
    // out <object>   
    var xmlOutput_PolicyReviewTaskNode =  writeXmlElement_Object(xmlOutput, xmlOutput_ImportNode, p_oPolicyDefinition, task_guid, "POLICYREVIEWTASK");
    
    //-----------------------------
    // out <meta>
    writeXmlElement_ObjectMetaData(xmlOutput, xmlOutput_PolicyReviewTaskNode, p_oPolicyDefinition);
    
    //-----------------------------
    // out <attr>s
    writeXmlElement_Attributes(xmlOutput, xmlOutput_PolicyReviewTaskNode, aAttrElements);
    // out change date <attr>
    writeXmlElement_ArisChangeDateAttribute(xmlOutput, xmlOutput_PolicyReviewTaskNode, p_oPolicyDefinition);
    
    //-----------------------------
    // out links as <attrRef>s
    writeXmlElement_Links(xmlOutput, xmlOutput_PolicyReviewTaskNode, aLinkElements, p_oPolicyDefinition);
    
    //===========================================
    //=== then export the policy definition ===
    
    var aObjList2CheckCycles = new Array();  // List with all children for checking cycles
	aObjList2CheckCycles.push(p_oPolicyDefinition);
	createXmlElement_StandardObject(xmlOutput, xmlOutput_ImportNode, xmlConfig_ObjectTypeMapping_pd, p_oPolicyDefinition, aObjList2CheckCycles, "POLICYDEFINITION");
    
    
    //===========================================================================
    //=== finally write the link from policy definition to policy review task ===
    
    //-----------------------------
    // create Link poining from original policy definition to this policy review task
    var xmlOutput_PolicyReviewTaskNode_PolicyDefinition = g_objDefXMLOutputObjectMap.get(p_oPolicyDefinition);
    aLinkElements = new Array();
    aLinkElements.push(new linkElement("policyreviewtask", "", null, null, task_guid));
    
    writeXmlElement_Links(xmlOutput, xmlOutput_PolicyReviewTaskNode_PolicyDefinition, aLinkElements, p_oPolicyDefinition);
    
    //update global map containing the reference from policy ObjDef to corresponding XMLOutput <object> tag of policy review task
    g_objDefXMLOutputObjectMap.put(p_oPolicyDefinition, xmlOutput_PolicyReviewTaskNode);
}

/*-------------------------------------------------------------------
    Looks up all ObjDefs assigned to a given ObjDef in a given Map. 
    Then creates linkElements using the given link ID.
---------------------------------------------------------------------*/
function createXmlElement_LinksByHashMap(p_oObjDef, p_linkElemId, p_oAssignmentHashMap) {
    
    var aLinkElements = new Array();
    
    var linkedObjs = p_oAssignmentHashMap.get(p_oObjDef);
    if (linkedObjs != null) {
        for (var i=0; i<linkedObjs.length; i++) {
            aLinkElements.push(new linkElement(p_linkElemId, "", null, linkedObjs[i], null));   
        }
    }

    return aLinkElements;    
}


/*---------------------------------------------------------------------------------------
    Create XML elements into output for a complete questionnaire template tree:
	- first all sections
	- then the questionnaire template
	If a cycle emerges during the recursion (check if object is already contained in
    list "aObjList2CheckCycles") an exception is thrown and the report aborts
---------------------------------------------------------------------------------------*/
function createXmlElement_QuestionnaireTemplate(xmlOutput, xmlOutput_ImportNode, hm_MappingObjects_Arcm2Aris, oObjDef, arcmType) {
	
	//first output sections
	var questionnaireTemplateRootGUID = oObjDef.GUID();
	var aObjList2CheckCycles = new Array();  // List with all children for checking cycles
	aObjList2CheckCycles.push(oObjDef);
	
    var xmlConfig_ObjectTypeMapping_qt = hm_MappingObjects_Arcm2Aris.get("QUESTIONNAIRE_TEMPLATE|OT_SURVEY_QUEST_TMPL");
    var xmlConfig_ObjectTypeMapping_section_cycleBreaker = hm_MappingObjects_Arcm2Aris.get("SECTION|OT_SURVEY_SECTION|cycleBreaker");
    var xmlConfig_ObjectTypeMapping_section = hm_MappingObjects_Arcm2Aris.get("SECTION|OT_SURVEY_SECTION");
    var xmlConfig_ObjectTypeMapping_question = hm_MappingObjects_Arcm2Aris.get("QUESTION|OT_SURVEY_QUESTION");
    var xmlConfig_ObjectTypeMapping_question_cycleBreaker = hm_MappingObjects_Arcm2Aris.get("QUESTION|OT_SURVEY_QUESTION|cycleBreaker");
    var oSortInfoQT2S = determineSortingMethodFromMappingObject(xmlConfig_ObjectTypeMapping_qt, "OT_SURVEY_SECTION");
    var oSortInfoS2S = determineSortingMethodFromMappingObject(xmlConfig_ObjectTypeMapping_section, "OT_SURVEY_SECTION");
    
    //determine the GUIDs to use for all sections during export - they may differ from the original ARIS ObjDef GUIDs
    var aSectionExportInfoResult = new Array();
    var aQuestionExportInfoResult = new Array();
    createSectionAndQuestionExportInfo(questionnaireTemplateRootGUID, aSectionExportInfoResult, aQuestionExportInfoResult);
    
    //export the found questions (both explicit and implicit) of this questionnaire template first
    //first as cycle breaker...
    for (var o=0; o<aQuestionExportInfoResult.length; o++) {
        createXmlElement_StandardObject(xmlOutput, xmlOutput_ImportNode, xmlConfig_ObjectTypeMapping_question_cycleBreaker, aQuestionExportInfoResult[o].exportObjDef, aObjList2CheckCycles, "QUESTION", 
                                        questionnaireTemplateRootGUID, 
                                        aQuestionExportInfoResult[o], aSectionExportInfoResult, aQuestionExportInfoResult);
    }
    //... then for real
    for (var n=0; n<aQuestionExportInfoResult.length; n++) {
        createXmlElement_StandardObject(xmlOutput, xmlOutput_ImportNode, xmlConfig_ObjectTypeMapping_question, aQuestionExportInfoResult[n].exportObjDef, aObjList2CheckCycles, "QUESTION", 
                                        questionnaireTemplateRootGUID, 
                                        aQuestionExportInfoResult[n], aSectionExportInfoResult, aQuestionExportInfoResult);
    }
    
    for (var m=0; m<aSectionExportInfoResult.length; m++) {
        //pass the questionnaire root GUID and output section entries as dummy cycle breaker 
        createXmlElement_StandardObject(xmlOutput, xmlOutput_ImportNode, xmlConfig_ObjectTypeMapping_section_cycleBreaker, aSectionExportInfoResult[m].exportObjDef, aObjList2CheckCycles, "SECTION", 
                                        questionnaireTemplateRootGUID, 
                                        aSectionExportInfoResult[m], aSectionExportInfoResult, aQuestionExportInfoResult);
    }
	for (var k=0; k<aSectionExportInfoResult.length; k++) {
        //pass the questionnaire root GUID and output sections with their standard info
		createXmlElement_StandardObject(xmlOutput, xmlOutput_ImportNode, xmlConfig_ObjectTypeMapping_section, aSectionExportInfoResult[k].exportObjDef, aObjList2CheckCycles, "SECTION", 
                                        questionnaireTemplateRootGUID, 
                                        aSectionExportInfoResult[k], aSectionExportInfoResult, aQuestionExportInfoResult);
    } 

	//then pass the questionnaire root GUID and output questionnaire template
	createXmlElement_StandardObject(xmlOutput, xmlOutput_ImportNode, xmlConfig_ObjectTypeMapping_qt, oObjDef, aObjList2CheckCycles, arcmType, questionnaireTemplateRootGUID, null, aSectionExportInfoResult);
}


/*---------------------------------------------------------------------------------------
    Create XML elements into output for a complete audit template tree:
	- first all audit step templates
	- then the audit template
	If a cycle emerges during the recursion (check if object is already contained in
    list "aObjList2CheckCycles") an exception is thrown and the report aborts
---------------------------------------------------------------------------------------*/
function createXmlElement_AuditTemplate(xmlOutput, xmlOutput_ImportNode, hm_MappingObjects_Arcm2Aris, oObjDef, arcmType) {

    //first output step templates
	var auditTemplateRootGUID = oObjDef.GUID();
	var aObjList2CheckCycles = new Array();  // List with all children for checking cycles
	aObjList2CheckCycles.push(oObjDef);
	
    var xmlConfig_ObjectTypeMapping_at = hm_MappingObjects_Arcm2Aris.get("AUDITTEMPLATE|OT_FUNC_INST");    
	var xmlConfig_ObjectTypeMapping_step = hm_MappingObjects_Arcm2Aris.get("AUDITSTEPTEMPLATE|OT_FUNC_INST");
    var oSortInfoAT2AST = determineSortingMethodFromMappingObject(xmlConfig_ObjectTypeMapping_at, "OT_FUNC_INST");
    var oSortInfoAST2AST = determineSortingMethodFromMappingObject(xmlConfig_ObjectTypeMapping_step, "OT_FUNC_INST");

    //get all direct and indirect steps by looking up the top steps first...
    var oCompleteStepTreeSet = new java.util.LinkedHashSet();
    var aLinkedTopSteps = g_auditTemplate2AuditStepTemplateHashMap.get(oObjDef); 
    if (aLinkedTopSteps != null) {
        //(sort them as defined in mapping.xml)
        aLinkedTopSteps = sortChildObjDefs(oSortInfoAT2AST.sortMethod, oSortInfoAT2AST.sortModelTypeNum, oObjDef, aLinkedTopSteps);
        //...then the subtree of steps for each top step (sorted)
		for (var j=0; j<aLinkedTopSteps.length; j++) { 
            getSubTree(aLinkedTopSteps[j], null, g_auditStepTemplate2subAuditStepTemplateHashMap, oCompleteStepTreeSet, oSortInfoAST2AST.sortMethod, oSortInfoAST2AST.sortModelTypeNum);
		} 
	}
    var aCompleteStepTree = convertHashSetToJSArray(oCompleteStepTreeSet);
    
    //reverse the sub tree to make sure we export the leaves first - this is needed for the export file consistency
    aCompleteStepTree.reverse();
    for (var k=0; k<aCompleteStepTree.length; k++) {
        //pass the audit template root GUID and output step templates
        createXmlElement_StandardObject(xmlOutput, xmlOutput_ImportNode, xmlConfig_ObjectTypeMapping_step, aCompleteStepTree[k], aObjList2CheckCycles, "AUDITSTEPTEMPLATE", auditTemplateRootGUID);
    }
	
	//then pass the audit template root GUID and output audit template
	createXmlElement_StandardObject(xmlOutput, xmlOutput_ImportNode, xmlConfig_ObjectTypeMapping_at, oObjDef, aObjList2CheckCycles, arcmType, auditTemplateRootGUID);
}


/*---------------------------------------------------------------------------------------
    Output of the current object with attributes and links
    - no multiple output of the same (ARIS) objects
    - recursion if connected objects were not output yet -> handle them first
      (under the same element Objects/obj_type)
    - If a cycle emerges during the recursion (check if object is already contained in
      list "aObjList2CheckCycles") an exception is thrown and the report aborts
---------------------------------------------------------------------------------------*/
function createXmlElement_StandardObject(xmlOutput, xmlOutput_ParentNode, xmlConfig_ObjectTypeMapping, oObjDef, aObjList2CheckCycles, arcmType, 
                                         rootGUID, 
                                         objDefExportInfo, aSectionExportInfoResult, aQuestionExportInfoResult) {
    
    //skip in debugging mode
    if (g_bDebug && xmlConfig_ObjectTypeMapping.getAttribute("info") != null && xmlConfig_ObjectTypeMapping.getAttribute("info").getValue() != "") {
        return;
    }
    
    rootGUID = rootGUID + ""; 
    
    //---- generate XML
    //special handling for SECTION and QUESTION: use the passed arcm GUID for export
    var guid = "";
    if (objDefExportInfo != null) {
        guid = objDefExportInfo.arcmGUID;
    } else {
        guid = oObjDef.GUID();
    }

    //-----------------------------------
    // --- Mapping of current object ---
    //-----------------------------------
    
    // --- mappingAttr -> Attr ---
    var aAttrElements = createXmlElement_Common_MappingAttributes(xmlConfig_ObjectTypeMapping, oObjDef, true, true, false);
    
    // --- mappingLink -> Link ---
    var aLinkElements = new Array();
    var xmlConfig_MapLinks = xmlConfig_ObjectTypeMapping.getChildren("mappingLink");
    var iterLink = xmlConfig_MapLinks.iterator();
    while (iterLink.hasNext()) {
        
        var xmlConfig_MapLink = iterLink.next();
        
		var linkApproachAttr = xmlConfig_MapLink.getAttribute("approach");
		if (linkApproachAttr != null && linkApproachAttr.getValue() != "" && linkApproachAttr.getValue() != g_approach) {
			continue;
		}
		
		var cxnTypeNum = xmlConfig_MapLink.getAttribute("aris_typenum").getValue();
        var cxnId = xmlConfig_MapLink.getAttribute("id").getValue();
        var bImplicitSections = xmlConfig_MapLink.getAttribute("implicitSections") != null && xmlConfig_MapLink.getAttribute("implicitSections").getValue() == "true";
        var bImplicitQuestions = xmlConfig_MapLink.getAttribute("implicitQuestions") != null && xmlConfig_MapLink.getAttribute("implicitQuestions").getValue() == "true";
        
        // SPECIAL CASE: Links to documents
        if (valuesAreEqual(cxnTypeNum, "DOCUMENT")) {
            aLinkElements = aLinkElements.concat( createXmlElement_Common_MappingLink_Document(oObjDef, cxnId) );  
        }
        // SPECIAL CASE: sections including implicit sections
        else if (bImplicitSections) {
            var aExplicitSectionLinkElements = createXmlElement_Common_MappingLink_Default(xmlConfig_MapLink, oObjDef, cxnId, cxnTypeNum, arcmType, rootGUID, aSectionExportInfoResult);
            aLinkElements = aLinkElements.concat( createXmlElement_SurveyManagement_MappingLink_AddDependingSections(oObjDef, rootGUID, cxnId, aExplicitSectionLinkElements, aSectionExportInfoResult) );
        }
        // SPECIAL CASE: questions including implicit questions
        else if (bImplicitQuestions) {
            var aExplicitQuestionLinkElements = createXmlElement_Common_MappingLink_Default(xmlConfig_MapLink, oObjDef, cxnId, cxnTypeNum, arcmType, rootGUID, aSectionExportInfoResult);
            aLinkElements = aLinkElements.concat( createXmlElement_SurveyManagement_MappingLink_AddDependingQuestions(oObjDef, rootGUID, cxnId, aExplicitQuestionLinkElements, aQuestionExportInfoResult) );
        }
        // DEFAULT:
        else {
            aLinkElements = aLinkElements.concat( createXmlElement_Common_MappingLink_Default(xmlConfig_MapLink, oObjDef, cxnId, cxnTypeNum, arcmType, rootGUID, aSectionExportInfoResult) );
        }
    }
    
    //for implicit sections and questions: add activating context links if activating questions are given
    if (objDefExportInfo != null && objDefExportInfo.aActivatingQuestionObjDefs != null) {
        var aActivatingQuestions = objDefExportInfo.aActivatingQuestionObjDefs;
        for (var i=0; i<aActivatingQuestions.length; i++) {
            aLinkElements.push(new linkElement("activatedInContextWith", "", null, aActivatingQuestions[i], null));
        }
        
    }
    
    //add root_guid to output
    if ((arcmType == "SECTION" || arcmType == "AUDITSTEPTEMPLATE")
        && rootGUID != null && rootGUID.length > 0) {          
        aAttrElements.push(new attrElement("root_guid", rootGUID));
    }
    
    //----------------------------------
    // --- Output of current object ---
    //----------------------------------
    // out <object>
    var xmlOutput_Object = writeXmlElement_Object(xmlOutput, xmlOutput_ParentNode, oObjDef, guid, arcmType);    
    
    //-----------------------------
    // out <meta>
    writeXmlElement_ObjectMetaData(xmlOutput, xmlOutput_Object, oObjDef);
    
    //-----------------------------
    // out <attr>s
    writeXmlElement_Attributes(xmlOutput, xmlOutput_Object, aAttrElements);
    // out change date <attr>
    writeXmlElement_ArisChangeDateAttribute(xmlOutput, xmlOutput_Object, oObjDef);

    //-----------------------------
    // out links as <attrRef>s
    writeXmlElement_Links(xmlOutput, xmlOutput_Object, aLinkElements, oObjDef);
    
    //-----------------------------
    // out <comment>
    if (xmlConfig_ObjectTypeMapping.getAttribute("info") != null) {
        var sInfoText = xmlConfig_ObjectTypeMapping.getAttribute("info").getValue();
        if (sInfoText == "cycleBreaker") {
            xmlOutput_Object.setAttribute("cyclebreak", "true");
        }
        var infoComment = new org.jdom.Comment(sInfoText);       
        var objectNodeContentList = xmlOutput_Object.getContent();
        objectNodeContentList.add(0, infoComment);
    }
    
    
    //update global map containing the reference from ObjDef to corresponding XMLOutput <object> tag
    g_objDefXMLOutputObjectMap.put(oObjDef, xmlOutput_Object);
}


/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
          ~~~~~~~ Functionality for creation of basic XML elements ~~~~~~~
                          (Attribute mapping, Link mapping) 
  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/

/*---------------------------------------------------------------------------------------
    Basic output method for reading attribute values based on the defined attribute
    mapping.
    Returns an JS-Array of "attrElement" structures.
 ----------------------------------------------------------------------------------------*/
function createXmlElement_Common_MappingAttributes(xmlConfig_ObjectTypeMapping, oObjDef, bLookupModel, bLookupRole, bSkipEmptyAttributes) {   
    
    var aAttrElements = new Array();                    
    var xmlConfig_MapAttrs = xmlConfig_ObjectTypeMapping.getChildren("mappingAttr");
    var iterAttr = xmlConfig_MapAttrs.iterator();
    
    // if model lookup is activated then check if there is either ""MODELGUID"" or "MODELNAME" amongst the attribute mappings 
    // - otherwise we don't need a model and can deactivae model lookup for saving time
    if (bLookupModel) {
        bLookupModel = false;
        while (!bLookupModel && iterAttr.hasNext()) {   
            var xmlConfig_MapAttr = iterAttr.next();
            var attrElemTypeNum 	= xmlConfig_MapAttr.getAttribute("aris_typenum").getValue();
            bLookupModel = bLookupModel || attrElemTypeNum == "MODELGUID" || attrElemTypeNum == "MODELNAME"; 
        }
    }    
    
    var oModel = null;          
    if (bLookupModel) {oModel = getModel(oObjDef);}             // get model
    var oRoleObjDef = null;
    if (bLookupRole) {oRoleObjDef = getRoleObject(oObjDef);}    // get role object (only relevant for OT_PERS_TYPE, for all other types null is returned)
        
    iterAttr = xmlConfig_MapAttrs.iterator();
    while (iterAttr.hasNext()) {   
        var xmlConfig_MapAttr = iterAttr.next();
		   
        var attrApproachAttr = xmlConfig_MapAttr.getAttribute("approach");
		if (attrApproachAttr != null && attrApproachAttr.getValue() != "" && attrApproachAttr.getValue() != g_approach) {
			continue;
		}
		
        var attrElemName 	 = xmlConfig_MapAttr.getAttribute("name").getValue();
        var attrElemType 	 = xmlConfig_MapAttr.getAttribute("type").getValue();
        var attrElemTypeNum  = xmlConfig_MapAttr.getAttribute("aris_typenum").getValue();
        var attrDefaultValue = null;
        if (xmlConfig_MapAttr.getAttribute("default_value") != null) {
            attrDefaultValue = xmlConfig_MapAttr.getAttribute("default_value").getValue();
        }
         
        var attrElemValue = getAttrValue(oObjDef, attrElemTypeNum, attrElemType, oModel, oRoleObjDef, attrDefaultValue);
        if (bSkipEmptyAttributes && attrElemValue == "") {
            continue;
        }

        aAttrElements.push(new attrElement(attrElemName, attrElemValue)); 
    }
    
    return aAttrElements;
}

/*---------------------------------------------------------------------------------------
    Looks up all document structures and checks which ones were created from the given
    ObjDef. Then links to their already exported XML objects are returned.
    Returns an JS-Array of "linkElement" structures.
 ----------------------------------------------------------------------------------------*/
function createXmlElement_Common_MappingLink_Document(p_oObjDef, cxnId) {

    var aLinkElements = new Array();
    
    for (var i = 0 ; i < m_aDocIdList.length ; i++ ) {
        if (p_oObjDef.IsEqual(m_aDocIdList[i].objDef)) {
            var linkElemId = cxnId;
            var docGUID = m_aDocIdList[i].docGUID;
            //documents never have connection attributes
            aLinkElements.push(new linkElement(linkElemId, "", null, null, docGUID));
        }
    }
    
    return aLinkElements;
}

/*---------------------------------------------------------------------------------------
    Checks which implict section dependencies are defined for the given section in
    combination with the given questionnaire template.
    Returns an JS-Array of "linkElement" structures.
 ----------------------------------------------------------------------------------------*/
function createXmlElement_SurveyManagement_MappingLink_AddDependingSections(p_oPlacingPositionObjDef, p_sQuestionnaireTemplateRootGUID, cxnId, aExplicitSectionLinkElements, aSectionExportInfoResult) {
  
    var aLinkElements = new Array();
    
    var oQT = ArisData.getActiveDatabase().FindGUID(p_sQuestionnaireTemplateRootGUID);   
    var aImplicitSectionPlacings = g_realImplicitSections.get(oQT);
       
    //stores the info which implicit question will be linked after which explicit question directly
    //(in case multiple explicit questions "share" the same implicit question then it is linked after the last of these explicit questions)
    var hm_implicitSectionExportInfo2ExplicitSection = new java.util.HashMap(); 
    
    //for each explicitly linked section:
    for (var h=0; h<aExplicitSectionLinkElements.length; h++) {
        var oCurrentExplicitLinkedSection = aExplicitSectionLinkElements[h].targetObjDef;
        //check if there are implicit sections for this explicit section
        //-> if so then add these links too
        for (var i=0; i<aImplicitSectionPlacings.length; i++) {
            for (var j=0; j<aSectionExportInfoResult.length; j++) {
                var positionAndSuperiorSectionMatch = (aImplicitSectionPlacings[i].oPositioningSection == null && aSectionExportInfoResult[j].superiorSectionObjDef == null)
                                                        || (aImplicitSectionPlacings[i].oPositioningSection != null && aSectionExportInfoResult[j].superiorSectionObjDef != null 
                                                            && aImplicitSectionPlacings[i].oPositioningSection.GUID() == aSectionExportInfoResult[j].superiorSectionObjDef.GUID());                                                   
                
                var targetSectionMatch = aImplicitSectionPlacings[i].oObjDefToPlace.GUID() == aSectionExportInfoResult[j].exportObjDef.GUID();
                //the implicit section shall be placed at this questionnaire template or section
                var implicitPlacingRelevant = ( (aImplicitSectionPlacings[i].oPositioningSection == null && p_sQuestionnaireTemplateRootGUID == p_oPlacingPositionObjDef.GUID())
                                                || (aImplicitSectionPlacings[i].oPositioningSection != null && aImplicitSectionPlacings[i].oPositioningSection.GUID() == p_oPlacingPositionObjDef.GUID())
                                              ); 
                var explicitSectionIsActivator = oCurrentExplicitLinkedSection.GUID() == aImplicitSectionPlacings[i].oActivatingSection.GUID(); //the explicit section contains the activator question of this dependent section
                
                if (positionAndSuperiorSectionMatch && targetSectionMatch 
                    && implicitPlacingRelevant && explicitSectionIsActivator) {
                    //store the export info for this implicit section and store (or overwrite) the related explicit section 
                    hm_implicitSectionExportInfo2ExplicitSection.put(aSectionExportInfoResult[j], oCurrentExplicitLinkedSection);    
                }
            }
        }
    }
   
    //for each explicitly linked section:
    for (var m=0; m<aExplicitSectionLinkElements.length; m++) {
        //add the link 
        aLinkElements.push(aExplicitSectionLinkElements[m]);
        var oExplicitSection = aExplicitSectionLinkElements[m].targetObjDef;
        //add all implicit questions links determined above
        for (var k=0; k<aSectionExportInfoResult.length; k++) {
            var oMappedExplicitSection = hm_implicitSectionExportInfo2ExplicitSection.get(aSectionExportInfoResult[k]);
            if (oMappedExplicitSection != null && oMappedExplicitSection.GUID() == oExplicitSection.GUID()) {
                aLinkElements.push(new linkElement(cxnId, "", null, null, aSectionExportInfoResult[k].arcmGUID));
            }
        }
    }

    return aLinkElements;
}

/*---------------------------------------------------------------------------------------
    Checks which implict question dependencies are defined for the given section in
    combination with the given questionnaire template.
    Returns an JS-Array of "linkElement" structures.
 ----------------------------------------------------------------------------------------*/
function createXmlElement_SurveyManagement_MappingLink_AddDependingQuestions(p_oPlacingPositionObjDef, p_sQuestionnaireTemplateRootGUID, cxnId, aExplicitQuestionLinkElements, aQuestionExportInfoResult) {

    var aLinkElements = new Array();
    
    var oQT = ArisData.getActiveDatabase().FindGUID(p_sQuestionnaireTemplateRootGUID);   
    var aImplicitQuestionPlacings = g_realImplicitQuestions.get(oQT);
    
    //stores the info which implicit question will be linked after which explicit question directly
    //(in case multiple explicit questions "share" the same implicit question then it is linked after the last of these explicit questions)
    var hm_implicitQuestionExportInfo2ExplicitQuestion = new java.util.HashMap(); 
    
    //for each explicitly linked question:
    for (var h=0; h<aExplicitQuestionLinkElements.length; h++) {
        var oCurrentExplicitLinkedQuestion = aExplicitQuestionLinkElements[h].targetObjDef;
        //check if there are implicit questions for this explicit question
        //-> if so then add these links too
        for (var i=0; i<aImplicitQuestionPlacings.length; i++) {
            for (var j=0; j<aQuestionExportInfoResult.length; j++) {
                var positionAndSuperiorSectionMatch = (aImplicitQuestionPlacings[i].oPositioningSection.GUID() == aQuestionExportInfoResult[j].superiorSectionObjDef.GUID());                                                   
                var targetQuestionMatch = aImplicitQuestionPlacings[i].oObjDefToPlace.GUID() == aQuestionExportInfoResult[j].exportObjDef.GUID();
                
                var implicitPlacingRelevant = aImplicitQuestionPlacings[i].oPositioningSection.GUID() == p_oPlacingPositionObjDef.GUID(); //the implicit question shall be placed at this section
                var explicitQuestionIsActivator = oCurrentExplicitLinkedQuestion.GUID() == aImplicitQuestionPlacings[i].oActivatingQuestion.GUID(); //the explicit question is the activator of this dependent question
                
                if (positionAndSuperiorSectionMatch && targetQuestionMatch 
                    && implicitPlacingRelevant && explicitQuestionIsActivator) {
                    //store the export info for this implicit question and store (or overwrite) the related explicit question 
                    hm_implicitQuestionExportInfo2ExplicitQuestion.put(aQuestionExportInfoResult[j], oCurrentExplicitLinkedQuestion);    
                }
            }
        }
    }
    
    //for each explicitly linked question:
    for (var m=0; m<aExplicitQuestionLinkElements.length; m++) {
        //add the link 
        aLinkElements.push(aExplicitQuestionLinkElements[m]);
        var oExplicitQuestion = aExplicitQuestionLinkElements[m].targetObjDef;
        //add all implicit questions links determined above
        for (var k=0; k<aQuestionExportInfoResult.length; k++) {
            var oMappedExplicitQuestion = hm_implicitQuestionExportInfo2ExplicitQuestion.get(aQuestionExportInfoResult[k]);
            if (oMappedExplicitQuestion != null && oMappedExplicitQuestion.GUID() == oExplicitQuestion.GUID()) {
                aLinkElements.push(new linkElement(cxnId, "", null, null, aQuestionExportInfoResult[k].arcmGUID));
            }
        }
    }

    return aLinkElements;
}

/*---------------------------------------------------------------------------------------
    Looks up all ObjDefs from the interal maps which fit to the the target type,
    connection type and direction. For each such object a link is created (and link 
    attributes if specified in the mapping file).
    Returns an JS-Array of "linkElement" structures.
 ----------------------------------------------------------------------------------------*/
function createXmlElement_Common_MappingLink_Default(xmlConfig_MapLink, p_oObjDef, cxnId, cxnTypeNum, arcmType, rootGUID, aSectionExportInfoResult) {

    var aLinkElements = new Array();
    
    var cxnTypeNum = getAttributeTypeNum(cxnTypeNum);
    //ignore mapping linke where the ARCM id is not set or the ARIS connection ID cannot be resolved
    if (isNaN(cxnTypeNum) || cxnTypeNum == -1 || cxnId == null || cxnId == "") {
        return aLinkElements;
    }
    
    //condition check - if p_oObjDef does not fulfill the link conditions then an empty Array is returned
    var xmlConfig_MapLinkConditions = xmlConfig_MapLink.getChildren("linkCondition");
    var iterConditions = xmlConfig_MapLinkConditions.iterator();
    while (iterConditions.hasNext()) {
        var condition = iterConditions.next();
        var conditionAttrTypeNumString = condition.getAttribute("aris_typenum").getValue();
        var attrType = condition.getAttribute("type").getValue();
        var expectedValueString = condition.getAttribute("value").getValue() + "";
        
        var objectAttrValue = getAttrValue_Default(p_oObjDef, conditionAttrTypeNumString, attrType);
        if ( objectAttrValue !== expectedValueString ) {
            return aLinkElements;
        }    
    }    
    
    var cxnTargetObjType = null;
    if (xmlConfig_MapLink.getAttribute("aris_target_objtype") != null) {
        cxnTargetObjType = getAttributeTypeNum( xmlConfig_MapLink.getAttribute("aris_target_objtype").getValue() );
    }
    var cxnTargetSymbolNum = null;
    if (xmlConfig_MapLink.getAttribute("aris_target_symbolnum") != null) {
        cxnTargetSymbolNum = getSymbolNum( xmlConfig_MapLink.getAttribute("aris_target_symbolnum").getValue() );
    }
    var groupRole = null;
    if (xmlConfig_MapLink.getAttribute("grouprole") != null) {
        groupRole = xmlConfig_MapLink.getAttribute("grouprole");
    }
    var sortMethod = null;
    if (xmlConfig_MapLink.getAttribute("sort_method") != null) {
        sortMethod = xmlConfig_MapLink.getAttribute("sort_method").getValue();
    }
    var sortModelTypeNum = null;
    if (xmlConfig_MapLink.getAttribute("sort_model_typenum") != null) {
        sortModelTypeNum = xmlConfig_MapLink.getAttribute("sort_model_typenum").getValue();
    }
    var bIsOutgoing = checkCxnDirection(xmlConfig_MapLink.getAttribute("direction"));
     
    if (cxnTypeNum < 1) {
        return aLinkElements;
    }
    
    var aChildObjDefs = new Array();

    aChildObjDefs = aChildObjDefs.concat( determineObjectsByLinkType(p_oObjDef, cxnId, cxnTypeNum, bIsOutgoing, cxnTargetObjType, cxnTargetSymbolNum, groupRole, arcmType));
    aChildObjDefs = sortChildObjDefs(sortMethod, sortModelTypeNum, p_oObjDef, aChildObjDefs);
     
    //create link elements
    for (var j = 0 ; j < aChildObjDefs.length ; j++ ) {
        
        var oChildDef = aChildObjDefs[j];      
        var linkElemId = cxnId;
        var linkDeact = isLinkDeactivated(p_oObjDef, oChildDef);
        
        //determine connection attributes   
        var aCxnAttrElements = new Array();
        if (xmlConfig_MapLink.getChildren("mappingAttr") != null && xmlConfig_MapLink.getChildren("mappingAttr").size() > 0) {
            var oCxn = getConnectionByObjDefs(p_oObjDef, oChildDef);
            if (oCxn != null) {
                aCxnAttrElements = createXmlElement_Common_MappingAttributes(xmlConfig_MapLink, oCxn, false, false, false);
            }
        }
        
        var sTargetGUID = oChildDef.GUID();     
        //special handling for SECTION: use the ARCM GUID from the calculated export info for this child section
        if (aSectionExportInfoResult != null) {        
            for (var k=0; k<aSectionExportInfoResult.length; k++) {          
                if (arcmType != "QUESTIONNAIRE_TEMPLATE" 
                    && oChildDef.GUID() == aSectionExportInfoResult[k].exportObjDef.GUID() && aSectionExportInfoResult[k].superiorSectionObjDef != null && p_oObjDef.GUID() == aSectionExportInfoResult[k].superiorSectionObjDef.GUID()) {
                    sTargetGUID = aSectionExportInfoResult[k].arcmGUID;
                }
                if (arcmType == "QUESTIONNAIRE_TEMPLATE" 
                    && oChildDef.GUID() == aSectionExportInfoResult[k].exportObjDef.GUID() && aSectionExportInfoResult[k].superiorSectionObjDef == null) {
                    sTargetGUID = aSectionExportInfoResult[k].arcmGUID;
                }
            }
        }  
        
        var newLinkElement = new linkElement(linkElemId, linkDeact, aCxnAttrElements, oChildDef, sTargetGUID);
        aLinkElements.push( newLinkElement );  
    }
    
    return aLinkElements;
}

/*---------------------------------------------------------------------------------------
    Returns the connection object between two given objects or null if they are not
    connected at all
---------------------------------------------------------------------------------------*/
function getConnectionByObjDefs(p_oFirstObjDef, p_oSecondObjDef) {
    var aCxns = p_oFirstObjDef.CxnList();
    for (var i=0; i<aCxns.length; i++) {
        if (   (aCxns[i].SourceObjDef().equals(p_oFirstObjDef) && aCxns[i].TargetObjDef().equals(p_oSecondObjDef) ) 
            || (aCxns[i].SourceObjDef().equals(p_oSecondObjDef) && aCxns[i].TargetObjDef().equals(p_oFirstObjDef) )
            ) {
            return aCxns[i]; 
        }
    } 
    return null;
}

/*---------------------------------------------------------------------------------------
    Checks the optional connection direction inside the mapping file.
    Default = outgoing connection -> returns true
    ingoing connection -> returns false
 ---------------------------------------------------------------------------------------*/        
function checkCxnDirection(p_direction) {
    if (p_direction != null && valuesAreEqual(p_direction.getValue(), "in")) {
        return false;
    }
    return true;
}

/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
         ~~~~~~~ Basic XML output functions common for each exported object ~~~~~~~
  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/

/*---------------------------------------------------------------------------------------
    Basic output method for the enclosing element <object> which is generated for each 
    exported object.
    The reference to the new created element <object> is returned.
 ----------------------------------------------------------------------------------------*/
function writeXmlElement_Object(xmlOutput, xmlOutput_Objects, oObjDef, guid, arcmType) {

    var xmlOutput_Object = xmlOutput.addElement(xmlOutput_Objects, "object");
    xmlOutput_Object.setAttribute("guid", guid);
    xmlOutput_Object.setAttribute("type", arcmType);
    
    if (g_bReadDeactivateAttribute) {
        var deactivated = getAttrValue(oObjDef, Constants.AT_DEACT, "Boolean", null, null, null);
        if (deactivated == "true") {
            xmlOutput_Object.setAttribute("purpose", "DEACTIVATE");
        }
    }
    
    return xmlOutput_Object;
}

/*---------------------------------------------------------------------------------------
    Basic output method for the <attr> elements of each exported object.
 ----------------------------------------------------------------------------------------*/
function writeXmlElement_Attributes(xmlOutput, xmlOutput_Object, aAttrElements) {
    for (var i = 0 ; i < aAttrElements.length ; i++ ) {    
        var xmlOutput_Attribute = xmlOutput.addElement(xmlOutput_Object, "attr");
        xmlOutput_Attribute.setAttribute("id", aAttrElements[i].name);
        xmlOutput_Attribute.setText(aAttrElements[i].value);
    } 
}

/*---------------------------------------------------------------------------------------
    Basic output method for the "aris_change_date" <attr> element.
 ----------------------------------------------------------------------------------------*/
function writeXmlElement_ArisChangeDateAttribute(xmlOutput, xmlOutput_Object, oObjDef) {
    var sChangeDate = getAttrValue(oObjDef, Constants.AT_LAST_CHNG_2, "Date", null, null, null);
    var xmlOutput_Attribute = xmlOutput.addElement(xmlOutput_Object, "attr");
    xmlOutput_Attribute.setAttribute("id", "aris_change_date");
    xmlOutput_Attribute.setText( sChangeDate );  
}

/*---------------------------------------------------------------------------------------
    Basic output method for the <attrRef> elements of each exported object.
 ----------------------------------------------------------------------------------------*/
function writeXmlElement_Links(xmlOutput, xmlOutput_Object, aLinkElements, oObjDef) {
    
    var hm_attrRefNodes = new Packages.java.util.HashMap();
    for (var i = 0 ; i < aLinkElements.length ; i++ ) {    
        var linkID = aLinkElements[i].linkId;
        //get or create <attrRef> element
        var xmlOutput_AttrRef = hm_attrRefNodes.get(linkID);
        if (xmlOutput_AttrRef == null) {
            var xmlOutput_AttrRef = xmlOutput.addElement(xmlOutput_Object, "attrRef");
            xmlOutput_AttrRef.setAttribute("id", aLinkElements[i].linkId);
            hm_attrRefNodes.put(linkID, xmlOutput_AttrRef);
        }
          
        //out <ref> info
        var xmlOutput_Ref = xmlOutput.addElement(xmlOutput_AttrRef, "ref");
        if (aLinkElements[i].targetGUID != null) {
            xmlOutput_Ref.setAttribute("guid", aLinkElements[i].targetGUID);
        } else {
            xmlOutput_Ref.setAttribute("guid", aLinkElements[i].targetObjDef.GUID());
        }   
        
        var sDeactFlag = new String(aLinkElements[i].deact);
        if (sDeactFlag.length > 0) {
            xmlOutput_Ref.setAttribute("purpose", "REMOVE");
        }
        
        //out link attributes
        var aCxnAttrElements = aLinkElements[i].aAttrElements;
        if (aCxnAttrElements != null) {
            //xmlOutput_Ref is now the enclosing "xmlOutput_Object"
            writeXmlElement_Attributes(xmlOutput, xmlOutput_Ref, aCxnAttrElements);
        }
    }
}


/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
         ~~~~~~~ Model mapping helper functions ~~~~~~~
  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/

/*---------------------------------------------------------------------------------------
    Determines 1st model where the given object has an occurrence
    - For objects of type "organisational unit" only organisational charts are considered
    - For objects of type "technical term" only technical term models are considered
    - For objects of type "function" only EPCs, VACDs, FADs, EBPMN collaboration and 
      EBPMN process diagrams are considered
      additionally it is checked if the function in the model is linked to a risk 
    - If no fitting model is found an arbitrary is returned
 ---------------------------------------------------------------------------------------*/
function getModel(p_objDef) {
    
    var oObjOccList = p_objDef.OccList();
    
    var aExportRelevantModels = new Array();
    var aNonExportRelevantModels = new Array();
    for (var i = 0 ; i < oObjOccList.length ; i++) {
        var oModel = oObjOccList[i].Model();
        if (isExportRelevant(oModel)) {
            aExportRelevantModels.push(oModel);
        }
        else {
            aNonExportRelevantModels.push(oModel);
        }
    }
    //sort the models by GUID - function 'sortByGUID' from ARCM-Utils.js
    aExportRelevantModels.sort(sortByGUID);
    aNonExportRelevantModels.sort(sortByGUID);
    
    var oModel = getBestModelCandidate(p_objDef, aExportRelevantModels);
    if (oModel == null) {
        oModel = getBestModelCandidate(p_objDef, aNonExportRelevantModels);
    }
    
    return oModel;
}

/*---------------------------------------------------------------------------------------
    Determines 1st model where the given object has an occurrence
    - For objects of type "organisational unit" only organisational charts are considered
    - For objects of type "technical term" only technical term models are considered
    - For objects of type "function" only EPCs, VACDs and FADs are considered
      additionally it is checked if the function in the model is linked to a risk 
    - If no fitting model is found an arbitrary is returned
 ---------------------------------------------------------------------------------------*/
function getBestModelCandidate(p_objDef, p_aModels) {    
    
    //Level 1: correct model type, for functions additional connections to risks
    for (var i = 0 ; i < p_aModels.length ; i++ ) {
        var oModel = p_aModels[i];
        switch(p_objDef.TypeNum()) {
            case Constants.OT_ORG_UNIT:
                if (oModel.TypeNum() == Constants.MT_ORG_CHRT) {
                    return oModel;
                }
                break;
            case Constants.OT_TECH_TRM:
                if (oModel.TypeNum() == Constants.MT_TECH_TRM_MDL) {
                    return oModel;
                }
                break;
            case Constants.OT_FUNC:
                if (oModel.TypeNum() == Constants.MT_EEPC || 
                    oModel.TypeNum() == Constants.MT_VAL_ADD_CHN_DGM ||
                    oModel.TypeNum() == Constants.MT_FUNC_ALLOC_DGM ||
                    oModel.TypeNum() == Constants.MT_ENTERPRISE_BPMN_COLLABORATION ||
                    oModel.TypeNum() == Constants.MT_ENTERPRISE_BPMN_PROCESS ) {
					return oModel;
                }
				//check for GDPR extension model "GDPR_PADD (Processing activity description diagram)"
                if (oModel.TypeNum() == getModelTypeNum("b0205e20-4aa5-11e7-43b7-08002721906d")) {
                    return oModel;
                }
                break;
        }
    }
    //Level 2: Functions - (only) correct model type
    if  (p_objDef.TypeNum() == Constants.OT_FUNC) {
        for (var i = 0 ; i < p_aModels.length ; i++ ) {
            var oModel = p_aModels[i];
            
            if (oModel.TypeNum() == Constants.MT_EEPC || 
                oModel.TypeNum() == Constants.MT_VAL_ADD_CHN_DGM ||
                oModel.TypeNum() == Constants.MT_FUNC_ALLOC_DGM) {
                    
                return oModel;
            }
        }
    }        
    //Level 3: first model with occurrence 
    if (p_aModels.length > 0) {
        return p_aModels[0];
    }
    //no model with occurrences
    return null;
}

/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
         ~~~~~~~ Find ObjDefs in internal lists/ maps - helper functions ~~~~~~~
  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
  
/*---------------------------------------------------------------------------------------
    Gets all objects from the given list that correspond to the given type (mandatory) 
    and the given symbol as default symbol (optional, can be null).
    Each object must be rated as qualified by method "isRelevantForExport()".
    Special handling for risk assessments.
 ---------------------------------------------------------------------------------------*/        
function getObjectsByType(p_objDefList, p_typeNum, p_symbolNum, p_arcmType) {
    var aObjDefs = new Array();
    
    //if p_arcmType is POLICYREVIEWTASK return the policy definitions they shall be generated from
    if (p_arcmType == "POLICYREVIEWTASK") {
        p_arcmType = "POLICYDEFINITION";
        p_typeNum = Constants.OT_POLICY;
    }
    
    //parse mapping typenums
    var aTypeNums = (p_typeNum + "").split(",");
    var oTypeNumsSet = new java.util.HashSet();
    for (var i=0; i<aTypeNums.length; i++) {
        var oTypeNum = getAttributeTypeNum(aTypeNums[i]);
        if (oTypeNum != null && oTypeNum != -1) { oTypeNumsSet.add( oTypeNum ); }
    }
    
    //parse mapping symbolnums
    var oSymbolNumsSet =  new java.util.HashSet();
    if (p_symbolNum != null) {
        var aSymbolNums = (p_symbolNum + "").split(",");
        for (var i=0; i<aSymbolNums.length; i++) {
            var oSymbolNum = getSymbolNum(aSymbolNums[i]);
            if (oSymbolNum != null && oSymbolNum != -1) { oSymbolNumsSet.add( oSymbolNum ); }
        }
    }
 
    for (var i = 0 ; i < p_objDefList.length ; i++ ) {       
        //in case the typeNum does not correspond to at least one of the mapping object
        if (!oTypeNumsSet.contains(p_objDefList[i].TypeNum())) {
            continue;
        }
        //in case the mapping object has defined symbols and symbolTypeNum does not correspond to at least one them
        if (!oSymbolNumsSet.isEmpty() && !oSymbolNumsSet.contains(p_objDefList[i].getDefaultSymbolNum())) {
            continue;
        }
        
        aObjDefs.push(p_objDefList[i]);
    }
    return aObjDefs;
}


/*---------------------------------------------------------------------------------------
    Determine linked objects by link type. Contains special handling for risks that 
    inherit linked objects from superior risks.
---------------------------------------------------------------------------------------*/
function determineObjectsByLinkType(oObjDef, cxnId, cxnTypeNum, isOutgoing, p_cxnTargetObjType, p_cxnTargetSymbolNum, grouprole, arcmType) {

    var aFoundObjects = new Array();
    
    //for risks only (inheritance): for assigned user groups return those already determined in ObjectClassification.js
    //Assignments of specific and generic risk are already aggregated -> if found, return immediately
    if (arcmType == "RISK") {
        //Risk manager
        if (cxnId == "manager_group") {
            var aFoundRiskManagers = g_risk2riskManagerHashMap.get(oObjDef);
            if (aFoundRiskManagers != null) {
                aFoundObjects = aFoundObjects.concat( aFoundRiskManagers );
            }
	        return aFoundObjects;
        }
        //Risk owner
        if (cxnId == "owner_group") {
            var aFoundRiskOwners = g_risk2riskOwnerHashMap.get(oObjDef);
            if (aFoundRiskOwners != null) {
                aFoundObjects = aFoundObjects.concat( aFoundRiskOwners );
            }
            return aFoundObjects;
        }
        //Risk reviewer
        if (cxnId == "reviewer_group") {       
            var aFoundRiskReviewers = g_risk2riskReviewerHashMap.get(oObjDef);
            if (aFoundRiskReviewers != null) {
                aFoundObjects = aFoundObjects.concat( aFoundRiskReviewers );
            }
	        return aFoundObjects;
        }    
    }
    
    if (arcmType == "AUDITTEMPLATE") {
        //for audit template -> top level audit step templates:
        //since *all* audit template steps have an ASSIGN connection to the audit template we must use the HashMap content here  
        if (cxnId == "steps") {
            var aFoundAuditStepTemplates = g_auditTemplate2AuditStepTemplateHashMap.get(oObjDef);
            if (aFoundAuditStepTemplates != null) {
                aFoundObjects = aFoundObjects.concat( aFoundAuditStepTemplates );
            }
	        return aFoundObjects;
        }
    }
    
    //"normal" links of risks and other objects
    aFoundObjects = getChildObjDefs(oObjDef, cxnTypeNum, isOutgoing, p_cxnTargetObjType, p_cxnTargetSymbolNum, grouprole, arcmType, cxnId); 
    
    //for risks only (inheritance): for each linktype below also inherit all objects of generic risk 
    //see ObjectClassification.js too see the inheritance algorithm which decides if the childern of specific or general risk are exported
    //Link types: 
    //  risk->function
    //  risk->ccounts
    //  risk->orgunits
    //  risk->appsys
    //  risk->category
    //  risk->risk manager
    //  risk->risk owner
    //  risk->risk reviewer
    if (aFoundObjects.length == 0 &&
        arcmType == "RISK" &&
        (   cxnId == "process" || cxnId == "finaccount" || cxnId == "orgunit" || cxnId == "appsystem" || cxnId == "risk_category" || cxnId == "manager_group" 
            || cxnId == "owner_group" || cxnId == "reviewer_group")
        ) {
        var oGenericRisk = g_specificRisk2genericRiskHashMap.get(oObjDef);
        if (oGenericRisk != null) {
            aFoundObjects = determineObjectsByLinkType(oGenericRisk, cxnTypeNum, cxnId, isOutgoing, p_cxnTargetObjType, p_cxnTargetSymbolNum, grouprole, arcmType);
        }
    }
      
    return aFoundObjects;
}

/*---------------------------------------------------------------------------------------
    Determines the child objects:
    - Source object of the incoming connection
    or
    - Destination object of the outgoing connection
    Restriction: Source / Target object is only relevant if it is contained in the list
    of selected objects.
 ---------------------------------------------------------------------------------------*/        
function getChildObjDefs(p_objDef, p_cxnTypeNum, p_isOutgoing, p_cxnTargetObjType, p_cxnTargetSymbolNum, p_groupRole, p_arcmType, p_cxnId) {
    var aChildObjDefs = new Array();
    var aCxns = new Array();

    if (p_isOutgoing) {
        var aCxns = p_objDef.CxnListFilter(Constants.EDGES_OUT | Constants.EDGES_ASSIGN, p_cxnTypeNum); // outgoing connections
    } else {
        var aCxns = p_objDef.CxnListFilter(Constants.EDGES_IN, p_cxnTypeNum);   // incoming connections
    }
        
    for (var i = 0 ; i < aCxns.length ; i++ ) {
        var oCxn = aCxns[i];
        var oChildDef;
        if (p_isOutgoing) {
            oChildDef = oCxn.TargetObjDef();    // target obj def
        } else {
            oChildDef = oCxn.SourceObjDef();    // source obj def
        }
        
        //if (optional) p_cxnTargetObjType matches with child's TypeNum (needed for risk's double used connection CT_AFFECTS to TechTerms and AppSysTypes)
        if ( (p_cxnTargetObjType == null || p_cxnTargetObjType == oChildDef.TypeNum())
             && (p_cxnTargetSymbolNum == null || p_cxnTargetSymbolNum == oChildDef.getDefaultSymbolNum()) 
            ) {
        
            if ( checkGroupRole(oChildDef, p_groupRole) 
                 && isObjectInList(oChildDef, g_oSelObjDefs)
                 && oChildDef.GUID() != p_objDef.GUID()
                ) {
                // only objects with correct role
                // only selected objects
                // only relevant objects
                // no cycle connection - ignore if ObjDef is connected to itself as child
                aChildObjDefs.push(oChildDef);  // target obj def 
            }
        }
    }
    
    return aChildObjDefs;
}

/*---------------------------------------------------------------------------------------
    Checks if the given Object is assigned to the correct role / role level
---------------------------------------------------------------------------------------*/
function checkGroupRole(p_objDef, p_groupRole) {
    if (p_groupRole == null) {
        // no Role/RoleLevel
        return true;
    } else {
        // Role necessary
        var roleDef = getRoleObject(p_objDef);
        return hasRoleAndLevel(roleDef, p_groupRole.getValue());
    }
    return false;
}

/*---------------------------------------------------------------------------------------
    Check if "p_oObjDef" is already contained in list/array "p_oObjDefList"
---------------------------------------------------------------------------------------*/
function isObjectInList(p_oObjDef, p_oObjDefList) {
    for (var i = 0; i < p_oObjDefList.length; i++) {
        if (p_oObjDef.equals(p_oObjDefList[i])) {
            return true;
        }
    }
    return false;
}

/*---------------------------------------------------------------------------------------
    Help function for determining a complete subtree of objects by a parent->child 
    HashMap.
    Used by:
    - createXmlElement_AuditTemplate (audit step templates)
    - createXmlElement_QuestionnaireTemplate (sections)
---------------------------------------------------------------------------------------*/
function getSubTree(p_parentObjDef, p_rootObjDef, p_oTreeHashMap, p_oResultSet, sSortMethod, sSortModelTypeNum) {

    //prevent infinity loop in case the section structure itself contains a loop
    if (!p_oResultSet.add(p_parentObjDef)) {
        return;
    }
    
    var aChildren = p_oTreeHashMap.get(p_parentObjDef);
    //sort the children by the given method before going into recursion
    aChildren = sortChildObjDefs(sSortMethod, sSortModelTypeNum, p_parentObjDef, aChildren);
    if (aChildren != null) {
        for (var i=0; i<aChildren.length; i++) {
            getSubTree(aChildren[i], p_rootObjDef, p_oTreeHashMap, p_oResultSet);
        }
    }
    
    //in case of survey management: after the explicit modeled section sub tree now add the implicit depending sections too
    if (p_rootObjDef != null) {   
        var qtImplicitSections = g_realImplicitSections.get(p_rootObjDef);
        if (qtImplicitSections != null) {
            for (var i=0; i<qtImplicitSections.length; i++) {
                p_oResultSet.add(qtImplicitSections[i].oObjDefToPlace);
            }
        }  
    }
}

/*---------------------------------------------------------------------------------------
    Returns an array of all given hierarchy ObjDefs which ensures that nowhere in it
    a parent hierarchy has a smaller index than its children hierarchies.
    This method is needed for the export file consistency where children have to 
    exported before the parents.
---------------------------------------------------------------------------------------*/
function sortHierarchiesChildrenFirst(p_aObjDefs) {   
    //determine top hierarchies
    var aTopHierarchies = new Array();
    for (var i=0; i<p_aObjDefs.length; i++) {
        var parentsSet = g_hierarchyChild2hierarchyParentsHashMap.get(p_aObjDefs[i]);
        if (parentsSet == null || parentsSet.isEmpty()) {
            aTopHierarchies.push(p_aObjDefs[i]);
        }
    }
    //fill an array by recursive tree parsing, starting with the top hierarchies
    var aTreeParsedHierarchies = new Array();
    for (var j=0; j<aTopHierarchies.length; j++) {
        parseHierarchyTreeTopDownRecursion(aTopHierarchies[j], aTreeParsedHierarchies);
    }   
    //revert the Array -> now all parent hierarchies always have a bigger index than their children
    aTreeParsedHierarchies.reverse();
    return aTreeParsedHierarchies;
}

/*---------------------------------------------------------------------------------------
    Returns an array of hierarchie elements by parsing their tree by the HashMap
    g_hierarchyParent2hierarchyChildrenHashMap.
---------------------------------------------------------------------------------------*/
function parseHierarchyTreeTopDownRecursion(p_objDef, aResult) {
    aResult.push(p_objDef);
    var childrenSet = g_hierarchyParent2hierarchyChildrenHashMap.get(p_objDef);
    if (childrenSet != null) {
        var iterObj = childrenSet.iterator(); 
        while (iterObj.hasNext()) {
            parseHierarchyTreeTopDownRecursion(iterObj.next(), aResult);
        }
    }
}

/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
             ~~~~~~~ Sorting helper functions ~~~~~~~
  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/ 

/*---------------------------------------------------------------------------------------
    Scans the <mappingLink> elements within the given <mappingObject> element. From the 
    first <mappingLink> containing the given target object type the sort method String
    is returned; "null" if no according <mappingLink> exists or no value is set.
---------------------------------------------------------------------------------------*/
function determineSortingMethodFromMappingObject(p_xmlConfig_ObjectTypeMapping, sTargetObjectType) {

    var xmlConfig_MapLinks = p_xmlConfig_ObjectTypeMapping.getChildren("mappingLink");
    for (var i=0; i<xmlConfig_MapLinks.size(); i++) {
        
		var xmlConfig_MapLink = xmlConfig_MapLinks.get(i);
		
		var linkApproachAttr = xmlConfig_MapLink.getAttribute("approach");
		if (linkApproachAttr != null && linkApproachAttr.getValue() != "" && linkApproachAttr.getValue() != g_approach) {
			continue;
		}
		
		var aris_target_objtype_attr = xmlConfig_MapLink.getAttribute("aris_target_objtype");
        if (aris_target_objtype_attr != null && aris_target_objtype_attr.getValue() == sTargetObjectType) {
            var sSortMethod = null;
            if (xmlConfig_MapLink.getAttribute("sort_method") != null) {
                sSortMethod = xmlConfig_MapLink.getAttribute("sort_method").getValue();
            }
            var sSortModelTypeNum = null;
            if (xmlConfig_MapLink.getAttribute("sort_model_typenum") != null) {
                sSortModelTypeNum = xmlConfig_MapLink.getAttribute("sort_model_typenum").getValue();
            }
            return new sortInfo(sSortMethod, sSortModelTypeNum);
        } 
    }
    return null;
}
  
/*---------------------------------------------------------------------------------------
    Sorts an JSArray of given child ObjDefs based on the given sort method string.
    If the sort method String is "null" then it does nothing.
    Depending on the sort method a model typenum or a parent ObjDef must be passed.
---------------------------------------------------------------------------------------*/
function sortChildObjDefs(sSortMethod, sSortModelTypeNum, p_oParentObjDef, p_aChildObjDefs) {
    
    var sortedChildObjDefs = p_aChildObjDefs;
    if (sSortMethod == "sort_by_xy_position") {
        var iModelTypeNum = getModelTypeNum( sSortModelTypeNum );
        sortedChildObjDefs = sortObjDefsByPositionInModel( p_aChildObjDefs, iModelTypeNum, "sortObjOccsByXYPos" );
    }
    
    if (sSortMethod == "sort_by_yx_position") {
        var iModelTypeNum = getModelTypeNum( sSortModelTypeNum );
        sortedChildObjDefs = sortObjDefsByPositionInModel( p_aChildObjDefs, iModelTypeNum, "sortObjOccsByYXPos" );
    }
     
    if (sSortMethod == "sort_by_occ_distance") {
        sortedChildObjDefs = sortObjDefsByDistanceOfLinkedOccs(p_oParentObjDef, p_aChildObjDefs);
    }
    
    return sortedChildObjDefs;
}

/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
         ~~~~~~~ Attribute value helper functions ~~~~~~~
  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/

/*---------------------------------------------------------------------------------------
    Determines the attribute value and returns it as JavaScript String.
    - The following keywords for parameter p_attrTypeNum are handled: 
        MODELGUID, 
        MODELNAME, 
        HIERARCHYTYPE, 
        OBJECTGUID, 
        ROLE, 
        ROLELEVEL,
		FULLNAME,
        FALSE, 
        TRUE,
        DATE_NOW, 
        ISMULTIPLE,
        REPORTPROPERTY#<property name>,
		ENDDATE#<start date attribute>#<duration attribute>,
		DURATION#<duration attribute>,
        CONSTANT#<constant value>
        EVENT_DRIVEN_ALLOWED#<standard boolean attribute>#<attr value attribute>#<attr value for 'event-driven'>
    
    - Standard attribute are handled by function "getAttrValue_Default"
---------------------------------------------------------------------------------------*/
function getAttrValue(p_objDef, p_attrTypeNum, p_sAttrType, p_model, p_roleObjDef, p_sAttrDefaultValue) {
    
    var sAttrValue = "";
    
    if (valuesAreEqual(p_attrTypeNum, "MODELGUID")) {sAttrValue = getGuid(p_model);
    
    } else if (valuesAreEqual(p_attrTypeNum, "MODELNAME")) {sAttrValue = getName(p_model);
    
    } else if (valuesAreEqual(p_attrTypeNum, "HIERARCHYTYPE")) {
        
        sAttrValue = "";
        if (p_objDef.TypeNum() == Constants.OT_FUNC && g_processSet.contains(p_objDef)) {sAttrValue = "process";}
        if (p_objDef.TypeNum() == Constants.OT_TECH_TRM) {sAttrValue = "financialaccount";}
        // OT_ORG_UNIT: org hierarchy or tester hierarchy
        if (p_objDef.TypeNum() == Constants.OT_ORG_UNIT) {
            sAttrValue = "";
            if (g_orgUnitSet.contains(p_objDef)) {sAttrValue = "orgunit";} 
            if (g_testerSet.contains(p_objDef)) {sAttrValue = "tester";}
        }  
        // OT_RISK_CATEGORY: Risk category hierarchy
        if (p_objDef.TypeNum() == Constants.OT_RISK_CATEGORY) {sAttrValue = "category";}
        // OT_APPL_SYS_TYPE: application system type hierarchy
        if (p_objDef.TypeNum() == Constants.OT_APPL_SYS_TYPE) {sAttrValue = "applicationsystemtype";}
        // OT_CLST: data cluster hierarchy
        if (p_objDef.TypeNum() == Constants.OT_CLST) {sAttrValue = "datacluster";}

    } else if (valuesAreEqual(p_attrTypeNum, "OBJECTGUID")) {sAttrValue = getGuid(p_objDef);
        
    } else if (valuesAreEqual(p_attrTypeNum, "ROLE")) {
        
        var localizedRoleName = "";
        if (p_roleObjDef != null) {       
            localizedRoleName = getRole(p_roleObjDef, g_nLoc);
            sAttrValue = getStaticMappingValue_Enum(localizedRoleName, p_sAttrType);
            if (sAttrValue == "") {
                localizedRoleName = getRole(p_roleObjDef, 1033);
                sAttrValue = getStaticMappingValue_Enum(localizedRoleName, p_sAttrType);
            }
        }
    
    } else if (valuesAreEqual(p_attrTypeNum, "ROLELEVEL")) {        
        
        if (p_roleObjDef != null) {
            sAttrValue = getName(p_roleObjDef).trim();
            var nIndex = sAttrValue.indexOf("_");
            if (nIndex >= 0) {
                sAttrValue = sAttrValue.substring(nIndex+1);
                sAttrValue = getStaticMappingValue_Enum(sAttrValue, p_sAttrType);
            }
        }
    
    } else if (valuesAreEqual(p_attrTypeNum, "FULLNAME")) {
	
		sAttrValue =  getAttrValue_Default(p_objDef, "AT_FIRST_NAME", "String")
					+ " "
					+ getAttrValue_Default(p_objDef, "AT_LAST_NAME", "String"); 
	
	} else if (valuesAreEqual(p_attrTypeNum, "FALSE")) {sAttrValue = "false"; 
        
    } else if (valuesAreEqual(p_attrTypeNum, "TRUE")) {sAttrValue = "true"; 
    
    } else if (valuesAreEqual(p_attrTypeNum, "DATE_NOW")) {sAttrValue = new Packages.java.util.Date().getTime();
        
    } else if (valuesAreEqual(p_attrTypeNum, "ISMULTIPLE")) {sAttrValue = getAttrValue_EnumMultiple(p_objDef, p_attrTypeNum, p_sAttrType);       
         
    } else if (java.lang.String(p_attrTypeNum).startsWith("REPORTPROPERTY#")) {  
	
		var propertyName = java.lang.String(p_attrTypeNum).substring("REPORTPROPERTY#".length);
		sAttrValue = Context.getProperty( propertyName );
	
	} else if (java.lang.String(p_attrTypeNum).startsWith("ENDDATE#")) {  
	
		var enddateCalculationParamString = java.lang.String(p_attrTypeNum).substring("ENDDATE#".length);
		sAttrValue = getAttrValue_EnddateCalculation(p_objDef, enddateCalculationParamString);
	
	} else if (java.lang.String(p_attrTypeNum).startsWith("DURATION#")) {
	
		var durationAttributeID = java.lang.String(p_attrTypeNum).substring("DURATION#".length);
		var sDurationStringValue = p_objDef.Attribute( eval("Constants." + durationAttributeID) , g_nLoc).GetValue(false);
        sAttrValue = getAttrValue_Duration(sDurationStringValue);
	
    } else if(java.lang.String(p_attrTypeNum).startsWith("FILESIZE#")){
        
        var filesizeAttrString = java.lang.String(p_attrTypeNum).substring("FILESIZE#".length);
          // reads file size using ADS link
        var nAttrTypeNum = getAttributeTypeNum(filesizeAttrString);
        if (nAttrTypeNum > 0) {
            if (valuesAreEqual(new String(filesizeAttrString).substring(0,12), "AT_ADS_LINK_")) {
                var adsComp = Context.getComponent("ADS");
                if (adsComp != null) {
                    var hyperLink = p_objDef.Attribute(nAttrTypeNum, g_nLoc).GetValue(true);
                    var document = adsComp.getDocumentByHyperlink(hyperLink);
                    if(document != null) {
                        var size = org.apache.commons.io.IOUtils.toByteArray(document.getDocumentContent()).length;
                        sAttrValue = size;
                    }
                }
            }
        }
    } else if (java.lang.String(p_attrTypeNum).startsWith("CONSTANT#")) {sAttrValue = java.lang.String(p_attrTypeNum).substring("CONSTANT#".length);

    } else if (java.lang.String(p_attrTypeNum).startsWith("EVENT_DRIVEN_ALLOWED#")) {
    
        var tokenList = p_attrTypeNum.split("#");
        if (tokenList.length == 4) {
            var eventDrivenBooleanAttributeTypeNumString = tokenList[1]; //the "normal" boolean attribute to export...
            var frequencyAttributeTypeNumString = tokenList[2]; //... except in those cases where this frequency attribute...
            var frequencyAttributeEventDrivenValueTypeNumString = tokenList[3]; //... has this value -> in this case we export "true"
            
            sAttrValue = getAttrValue_eventDrivenAllowed(p_objDef, p_sAttrType, eventDrivenBooleanAttributeTypeNumString, frequencyAttributeTypeNumString, frequencyAttributeEventDrivenValueTypeNumString);
        }
    
	} else { 
        sAttrValue = getAttrValue_Default(p_objDef, p_attrTypeNum, p_sAttrType);
    }
    
    if ((sAttrValue == null || sAttrValue == "") && p_sAttrDefaultValue != null) {
        sAttrValue = p_sAttrDefaultValue;
    }
    
    return sAttrValue += "";
}

/*---------------------------------------------------------------------------------------
    Determines the attribute value of standard attributes.
    The following attribute types are handled specially: "Date", "Boolean" "Enum<XXX>"     
---------------------------------------------------------------------------------------*/
function getAttrValue_Default(p_objDef, p_attrTypeNum, p_sAttrType) {
    
    // Enumeration
    if (valuesAreEqual(new String(p_sAttrType).substring(0,4), "Enum") && p_attrTypeNum != null) {
        
        //check if preselected value shall be used for XML export
        var aAttrTypeNumParts = (p_attrTypeNum + "").split("#");
        //no preselection specified
        if (aAttrTypeNumParts.length == 1) {
            var enumValue = getAttrValue_Enum(p_objDef, p_attrTypeNum, p_sAttrType);
            return enumValue;
        }
        //preselection specified
        if (aAttrTypeNumParts.length == 2) {
            return aAttrTypeNumParts[1];
        }
        return "";
    }
    
    //Non-Enumeration
    var nAttrTypeNum = getAttributeTypeNum(p_attrTypeNum);
    if (nAttrTypeNum > 0) {
        
        // Date
        if (valuesAreEqual(p_sAttrType, "Date")) {
            if (p_objDef.Attribute(nAttrTypeNum, g_nLoc).MeasureValue() == null) {
                return "";
            }
            return p_objDef.Attribute(nAttrTypeNum, g_nLoc).MeasureValue().getTime();
        } 
        // Boolean
        if (valuesAreEqual(p_sAttrType, "Boolean")) {
            if (isboolattributetrue(p_objDef, nAttrTypeNum, g_nLoc)) {
                return "true";
            }
            return "false";
        }
        // Amount of currency (not currency type!)
        if (valuesAreEqual(p_sAttrType, "CurrencyAmount")) {
            var value = p_objDef.Attribute(nAttrTypeNum, g_nLoc);
            if (value.IsMaintained()){
                return value.MeasureValue(true);
            } else {
                return "";
            }
        }

        // converts ADS link to GUID
        if (valuesAreEqual(new String(p_attrTypeNum).substring(0,12), "AT_ADS_LINK_") && p_attrTypeNum != null) {
            return ARCM.transformADSLinkToGUID(p_objDef.Attribute(nAttrTypeNum, g_nLoc).GetValue(true));
        }
        
        // String
        if (valuesAreEqual(p_sAttrType, "String")) {
            return p_objDef.Attribute(nAttrTypeNum, g_nLoc).GetValue(true);
		}
        
        // default
        return p_objDef.Attribute(nAttrTypeNum, g_nLoc).GetValue(false);
        
    }
    // no valid attr type num
    return "";
}

/*---------------------------------------------------------------------------------------
    Determines the value from the enum mapping corresponding to the given value.
---------------------------------------------------------------------------------------*/
function getStaticMappingValue_Enum(p_value, p_sAttrType) {
    if (valuesAreEqual(new String(p_sAttrType).substring(0,4), "Enum")) {
        
        var xmlConfig_MappingEnumElement = getXmlConfigMappingEnumElement(p_sAttrType);
        if (xmlConfig_MappingEnumElement != null) {

            var xmlConfig_EnumItems = xmlConfig_MappingEnumElement.getChildren("enumItem");
        
            var iterItems = xmlConfig_EnumItems.iterator();
            while (iterItems.hasNext()) {
                var xmlConfig_EnumItem = iterItems.next();
                
                var sTypeNum = xmlConfig_EnumItem.getAttribute("aris_typenum").getValue();
                if (valuesAreEqual(sTypeNum, p_value)) {
                    return xmlConfig_EnumItem.getAttribute("id").getValue();                
                }
            }
        }
        // entry not found in enum mapping
        return "";
    }
    return p_value;
}

/*---------------------------------------------------------------------------------------
    Determines the value of enum attributes at the given ObjDef by reading the selected
    Enum value at the ObjDef and comparing it to the enum item mappings from the
    mapping XML.
---------------------------------------------------------------------------------------*/
function getAttrValue_Enum(p_objDef, p_attrTypeNum, p_sAttrType) {
    
    var xmlConfig_MappingEnumElement = getXmlConfigMappingEnumElement(p_sAttrType);
    if (xmlConfig_MappingEnumElement == null) {
        return "";
    }                
    
    var nAttrTypeNum = getAttributeTypeNum(p_attrTypeNum);
    if (nAttrTypeNum > 0) {    
        var oAttribute = p_objDef.Attribute(nAttrTypeNum, g_nLoc);
        var sUnitTypeNum = "";
        if (oAttribute.IsMaintained()) {
            sUnitTypeNum = oAttribute.MeasureUnitTypeNum();
        }
        
        var sAttrGUID = null;
        if (ArisData.ActiveFilter().isUserDefinedAttrType(nAttrTypeNum)) {
            sAttrGUID = ArisData.ActiveFilter().UserDefinedAttributeTypeGUID(nAttrTypeNum);
        }

        var enumValue = getXmlEnumValueViaMapping(xmlConfig_MappingEnumElement, sUnitTypeNum, sAttrGUID);
        return enumValue;
    }
    // no valid attr type num
    return "";        
}

/*---------------------------------------------------------------------------------------
    Determines the value of the XML enum element whose attribute "aris_typenum" fits to
    the given value "p_sUnitTypeNum" (including processing with "getAttributeTypeNum()")
    The value of "aris_typenum" from XML can have two meanings:
    1) A reference to a boolean attribute (for items of multiple enums - Example: assertion)
    2) A reference to a attribute value type (for items of single enums - Example: impact)
    In each case the value of "aris_typenum" can be a direct type number, a 
    constant reference or a GUID.
    
    Parameter p_attributeValueTypeGUID is only needed when an user defined attribute value
    type of an user defined value attribute must be resolved by both attribute and
    attribute value type GUID.
---------------------------------------------------------------------------------------*/    
function getXmlEnumValueViaMapping(p_xmlConfig_MappingEnumElement, p_sUnitTypeNum, p_attributeValueTypeGUID) {
    
    var xmlConfig_MappingEnumItems = p_xmlConfig_MappingEnumElement.getChildren("enumItem");
    var iterItems = xmlConfig_MappingEnumItems.iterator();
    
    while (iterItems.hasNext()) {
        var xmlConfig_MappingEnumItem = iterItems.next();
        var sXMLTypeNumValue = xmlConfig_MappingEnumItem.getAttribute("aris_typenum").getValue();
        
        //case 1) - Boolean attribute as part of a multi enum mapping
        var attrValueTypeNum = getAttributeTypeNum_Enum(sXMLTypeNumValue);
        
        //case 2) - reference for a value type of a value attribute
        //maybe special case 2) where the enumItem aris_type_num points at an user-defined attribute value type of an 
        //user-defined value attribute
        if (attrValueTypeNum == undefined && p_attributeValueTypeGUID != null && p_attributeValueTypeGUID != ""
            && (sXMLTypeNumValue != null && sXMLTypeNumValue.length() > 0)
            ) {
            attrValueTypeNum = ArisData.ActiveFilter().UserDefinedAttributeValueTypeNum(p_attributeValueTypeGUID, sXMLTypeNumValue);        
        }
        
        if ( (attrValueTypeNum != undefined) && valuesAreEqual(attrValueTypeNum, p_sUnitTypeNum) ) {   
            return xmlConfig_MappingEnumItem.getAttribute("id").getValue();                
        }
    }
    // entry not found in enum mapping    
    return "";
}

/*---------------------------------------------------------------------------------------
    Determines the value of "multiple" enums. The values are separated by commas.
---------------------------------------------------------------------------------------*/        
function getAttrValue_EnumMultiple(p_objDef, p_attrTypeNum, p_sAttrType) {
    var xmlConfig_MappingEnumElement = getXmlConfigMappingEnumElement(p_sAttrType);
    if (xmlConfig_MappingEnumElement == null) {
        return "";
    }                
    var sAttrValue = "";
    
    var xmlConfig_EnumItems = xmlConfig_MappingEnumElement.getChildren("enumItem");
    var iterItems = xmlConfig_EnumItems.iterator();
    
    while (iterItems.hasNext()) {
        var xmlConfig_EnumItem = iterItems.next();
        var nAttrTypeNum = getAttributeTypeNum(xmlConfig_EnumItem.getAttribute("aris_typenum").getValue());
        if (nAttrTypeNum > 0) {     
            if (isboolattributetrue(p_objDef, nAttrTypeNum, g_nLoc)) {
                if (sAttrValue.length > 0) {
                    sAttrValue = sAttrValue.concat(",");
                }
                sAttrValue = sAttrValue.concat(xmlConfig_EnumItem.getAttribute("id").getValue());                
            }
        }
    }
    return sAttrValue;
}

/*----------------------------------------------------------------------------------------------
    Determines the element inside the enum mapping by means of a partial String.
    Example: "EnumsomeID" 	-> EnumElement "mappingEnum" with attribute "aris_enum" = "someID".
			 not found?		-> EnumElement "mappingEnum" with attribute "enum" = "someID".
------------------------------------------------------------------------------------------------*/
function getXmlConfigMappingEnumElement(p_sAttrType){
    var sEnum = new String(p_sAttrType).substring(4); 
    var xmlConfig_MappingEnumElements = xmlConfig_Mapping.getChild("mappingEnums").getChildren("mappingEnum");
    
    var iterEnum = xmlConfig_MappingEnumElements.iterator();
    while (iterEnum.hasNext()) {
        var xmlConfig_MappingEnumElement = iterEnum.next();
        if ( xmlConfig_MappingEnumElement.getAttribute("aris_enum") != null
             && valuesAreEqual(xmlConfig_MappingEnumElement.getAttribute("aris_enum").getValue(), sEnum) ) {
            return xmlConfig_MappingEnumElement;
        }
        if ( xmlConfig_MappingEnumElement.getAttribute("enum") != null
             && valuesAreEqual(xmlConfig_MappingEnumElement.getAttribute("enum").getValue(), sEnum) ) {
            return xmlConfig_MappingEnumElement;
        }
    }
    return null;
}

/*---------------------------------------------------------------------------------------
    Reads the String attribute value of a duration attribute and converts it in
    timeInMillis since 1970 as String.
---------------------------------------------------------------------------------------*/
function getAttrValue_Duration(sDurationValue) {

    var mults = new Array(1, 60, 3600, 86400);
    
    sDurationValue = sDurationValue.trim();
    if (sDurationValue == null || sDurationValue.length() == 0) {return 0;}
    
    var timeInMillis = 0;
    var tokenList = java.util.Arrays.asList(sDurationValue.split(":"));
    if(tokenList.isEmpty()) {return 0;}
    java.util.Collections.reverse(tokenList);//seconds always first

    for (var i=0; i<Math.min(tokenList.size(), 4); i++) {
        timeInMillis += java.lang.Long.parseLong(tokenList.get(i)) * mults[i];
    }

    return new java.lang.String(timeInMillis * 1000);    
}

/*---------------------------------------------------------------------------------------
    Checks if the given frequency attribute is maintained and has the given value
    indicating "Event-driven".
    If so then "true" is returned, otherwise the standard boolean value read from
    the given boolean is returned.
---------------------------------------------------------------------------------------*/
function getAttrValue_eventDrivenAllowed(   p_objDef,
                                            p_sAttrType, 
                                            eventDrivenBooleanAttributeTypeNumString, 
                                            frequencyAttributeTypeNumString, 
                                            frequencyAttributeEventDrivenValueTypeNumString ) {
    
    //which value has the frequency enumeration set?
    var frequencyAttributeTypeNum = getAttributeTypeNum(frequencyAttributeTypeNumString);
    var frequencyAttributeMaintainedValueTypeNum = null;
    if (frequencyAttributeTypeNum > 0) {    
        var frequencyAttribute = p_objDef.Attribute(frequencyAttributeTypeNum, g_nLoc);
        if (frequencyAttribute.IsMaintained()) {
            frequencyAttributeMaintainedValueTypeNum = frequencyAttribute.MeasureUnitTypeNum();
        }
    }
    //which value would be the special event driven value?
    var frequencyAttributeEventDrivenValueTypeNum = getAttributeTypeNum(frequencyAttributeEventDrivenValueTypeNumString);
    
    //if they are equal (i.e. event-driven is set) then return "true"
    if (frequencyAttributeMaintainedValueTypeNum + "" == frequencyAttributeEventDrivenValueTypeNum + "") {
        return "true";
    }
    //otherwise return the boolean value of the event driven boolean attribute as usual
    else {
        //p_sAttrType is "Boolean"
        return getAttrValue_Default(p_objDef, eventDrivenBooleanAttributeTypeNumString, p_sAttrType);
    }    
        
}

/*---------------------------------------------------------------------------------------
    Calls ObjDef.getEndDate(int startATN, int durationATN, int localeID).
    This method then calculates the enddate based on the attribute values and the
    boolean flag AT_WEEKEND_OFF.
---------------------------------------------------------------------------------------*/
function getAttrValue_EnddateCalculation(p_objDef, enddateCalculationParamString) {
    
    var tokenList = java.util.Arrays.asList(enddateCalculationParamString.split("#"));
    var startdateAttributeTypeNum = eval( "Constants." + tokenList.get(0) );
    var durationAttributeTypeNum = eval( "Constants." + tokenList.get(1) );
    
    var enddate = p_objDef.getEndDate(startdateAttributeTypeNum, durationAttributeTypeNum, g_nLoc);
    if (enddate == null) {return null;}
    return enddate.getTime();
}

/*---------------------------------------------------------------------------------------
    Checks if both Strings are equal.
    -> returns true, else false
 ---------------------------------------------------------------------------------------*/        
function valuesAreEqual(p_value1, p_value2) {
    if (p_value1 != null && p_value2 != null) {
        var sValue1 = java.lang.String(p_value1);
        var sValue2 = java.lang.String(p_value2);        
        return (sValue1.compareToIgnoreCase(sValue2) == 0);
    }
    return false;
}


/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
         ~~~~~~~ Common helper functions ~~~~~~~
  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/

/*---------------------------------------------------------------------------------------
    Determines the GUID.
 ---------------------------------------------------------------------------------------*/        
function getGuid(p_item) {
    if (p_item == null) {return "";}
    return p_item.GUID();
}

/*---------------------------------------------------------------------------------------
    Determines the name by attribute "AT_NAME".
 ---------------------------------------------------------------------------------------*/        
function getName(p_item) {
    if (p_item == null) {
        return "";
    }      
    return p_item.Attribute(Constants.AT_NAME, g_nLoc).GetValue(true);
}

/*---------------------------------------------------------------------------------------
    Determines the role ID by attribute "AT_NAME" and given Locale
 ---------------------------------------------------------------------------------------*/        
function getRole(p_item, p_nLocale) {
    if (p_item == null) {
        return "";
    }
    sAttrValue = p_item.Attribute(Constants.AT_NAME, p_nLocale).GetValue(true).trim();
    if (sAttrValue == null) {
        return "";
    }
    
    var nIndex = sAttrValue.indexOf("_");
    if (nIndex >= 0) {
        return sAttrValue.substring(0, nIndex);
    }  
    return "";
}

/*---------------------------------------------------------------------------------------
    Determines if the given ObjDef has the same role and role level String as name.
    First checks the attribute "AT_NAME" with the executing locale, then with the US
    english locale.
    Returns true if at least one attribute value matches, otherwise false.
 ---------------------------------------------------------------------------------------*/        
function hasRoleAndLevel(p_item, p_sRoleLevelString) {
    if (p_item == null || p_sRoleLevelString == null || p_sRoleLevelString == "") {
        return false;
    }
    var itemRoleLevelString = p_item.Attribute(Constants.AT_NAME, g_nLoc).GetValue(true);
    if (itemRoleLevelString == p_sRoleLevelString) {
        return true;
    }
    var itemRoleLevelString = p_item.Attribute(Constants.AT_NAME, 1033).GetValue(true);
    if (itemRoleLevelString == p_sRoleLevelString) {
        return true;
    }
    return false;
}


/*---------------------------------------------------------------------------------------
    Create a new GUID.
 ---------------------------------------------------------------------------------------*/        
function createGuid() {
    return Packages.com.aris.modeling.server.common.AGUIDGenerator.theInstance().createNewGUID().toExtendedString();
}

/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
       ~~~~~~~ ARCM GUID cache lookup functions ~~~~~~~
  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
 
/*------------------------------------------------------------------------------------------------------------------------------
    Returns the ARCM GUID of the document AppObject corresponding to the given name and link values, 
    assigned at the master AppObject defined by the master ObjDef GUID.
    Used for DOCUMENT.
 ------------------------------------------------------------------------------------------------------------------------------*/ 
function getDocumentGUIDFromGUIDMappingCache(sDocName, sDocLink, sMasterObjDefGUID) {
    return hmArisGUID2ArcmGUID.get(sDocName + "|" + sDocLink + "|" + sMasterObjDefGUID + "|" + g_arcmEnvironmentID);
}
/*------------------------------------------------------------------------------------------------------------------------------
    Stores the ARCM GUID of the document AppObject corresponding to the given name and link values, 
    assigned at the master AppObject defined by the master ObjDef GUID.
    Used for DOCUMENT.
 ------------------------------------------------------------------------------------------------------------------------------*/ 
function storeDocumentGUIDInGUIDMappingCache(sDocName, sDocLink, sMasterObjDefGUID, sARCMGUID) {
    hmArisGUID2ArcmGUID.put(sDocName + "|" + sDocLink + "|" + sMasterObjDefGUID + "|" + g_arcmEnvironmentID, sARCMGUID);
}
/*------------------------------------------------------------------------------------------------------------------------------
    Returns the ARCM GUID of the AppObject corresponding to the given master ObjDef GUID.
    Used for POLICYREVIEWTASK.
 ------------------------------------------------------------------------------------------------------------------------------*/ 
function getPolicyReviewTaskGUIDFromGUIDMappingCache(sMasterObjDefGUID) {
    return hmArisGUID2ArcmGUID.get(sMasterObjDefGUID + "|" + g_arcmEnvironmentID);
}
/*------------------------------------------------------------------------------------------------------------------------------
    Stores the ARCM GUID of the AppObject corresponding to the given master ObjDef GUID.
    Used for POLICYREVIEWTASK.
 ------------------------------------------------------------------------------------------------------------------------------*/
function storePolicyReviewTaskGUIDInGUIDMappingCache(sMasterObjDefGUID, sARCMGUID) {
    hmArisGUID2ArcmGUID.put(sMasterObjDefGUID + "|" + g_arcmEnvironmentID, sARCMGUID);
}

/*------------------------------------------------------------------------------------------------------------------------------
    Returns the the loaded AppObject based on the given GUID.
 ------------------------------------------------------------------------------------------------------------------------------*/ 
function getFromObjDefMappingCache(sObjDefGUID) {
    return hmObjDefGUID2ARCMAppObject.get(sObjDefGUID + "|" + g_arcmEnvironmentID);
}
/*------------------------------------------------------------------------------------------------------------------------------
    Stores the the loaded AppObject based on the given GUID.
 ------------------------------------------------------------------------------------------------------------------------------*/ 
function storeInObjDefMappingCache(sObjDefGUID, oARCMAppObject) {
    hmObjDefGUID2ARCMAppObject.put(sObjDefGUID + "|" + g_arcmEnvironmentID, oARCMAppObject);
}

/*------------------------------------------------------------------------------------------------------------------------------
    Returns an JS-Array of the the loaded sub AppObject based on the given master AppObject.
 ------------------------------------------------------------------------------------------------------------------------------*/
function getFromMasterObjectCache(oMasterARCMAppObject) {
    var aSubObjects = hmMasterObject2SubObjects.get(oMasterARCMAppObject);
    if (aSubObjects == null) {
        aSubObjects = new Packages.java.util.HashSet();
        hmMasterObject2SubObjects.put(oMasterARCMAppObject, aSubObjects);
    }
    return convertHashSetToJSArray(aSubObjects);
}
/*------------------------------------------------------------------------------------------------------------------------------
    Store a single loaded sub AppObject based on the given master AppObject.
 ------------------------------------------------------------------------------------------------------------------------------*/
function storeInMasterObjectCache(oMasterARCMAppObject, oSubARCMAppObject) {
    var aSubObjects = hmMasterObject2SubObjects.get(oMasterARCMAppObject);
    if (aSubObjects == null) {
        aSubObjects = new Packages.java.util.HashSet();
        hmMasterObject2SubObjects.put(oMasterARCMAppObject, aSubObjects);
    }
    aSubObjects.add(oSubARCMAppObject);
}

/*----------------------------------------------------------------------------------------------------------------------------------------------------
    Due to depending sections the same section ObjDef from ARIS can now represent multiple (independent) section AppObjs in the same 
    questionnaire template in ARCM. 
    The section export info is calculated to determine which ARCM GUIDs shall be used when exporting ARIS ObjDef sections multiple times
    for the same questionnaire template.
 -----------------------------------------------------------------------------------------------------------------------------------------------------*/
function createSectionAndQuestionExportInfo(qtGUID, aSectionExportInfoResult, aQuestionExportInfoResult) {
    
    var oGUIDRequestResultList = new java.util.ArrayList();
    
    //the QT on ARCM side
    if (g_arcmEnvironmentID != null) {
        oGUIDRequestResultList = ARCM.findByGUID(qtGUID, g_arcmEnvironmentID, g_nLoc);
    }
    var qtAppObject = null;
    if (!oGUIDRequestResultList.isEmpty()) {
        qtAppObject = oGUIDRequestResultList.get(0);
    }  
    //the top sections on ARCM side
    var oARCMSections = new java.util.ArrayList();
    if (qtAppObject != null) {
        oARCMSections.addAll( qtAppObject.getListAttribute("sections").getConnectedObjects() );
    }
    
    //the QT on ARIS side
    var qtObjDef = ArisData.getActiveDatabase().FindGUID(qtGUID);
    //the top sections on ARIS side - explicit 
    var aSections = new Array();
    aSections = aSections.concat(g_questionnaireTemplate2sectionsHashMap.get(qtObjDef));
    //the top sections on ARIS side - implicit
    var aAllImplicitSections = g_realImplicitSections.get(qtObjDef);
    var aImplicitSections = new Array();
    for (var i=0; i<aAllImplicitSections.length; i++) {
        if (aAllImplicitSections[i].oPositioningSection == null) {
            aImplicitSections.push(aAllImplicitSections[i].oObjDefToPlace);
        }
    }
    //all implicit questions of this questionnaire template
    var aAllImplicitQuestions = new Array();
    if (g_realImplicitQuestions.get(qtObjDef) != null) {
        aAllImplicitQuestions = g_realImplicitQuestions.get(qtObjDef);
    }
    
    //recursive comparison between ARCM section tree and ARIS section tree
    var aARCMSections = convertHashSetToJSArray(oARCMSections);
    var aARCMQuestions = new Array();
    createSectionAndQuestionExportInfoRecursive(null, aSections, aAllImplicitSections, aAllImplicitQuestions, aARCMSections, aARCMQuestions, aSectionExportInfoResult, aQuestionExportInfoResult);
}

/*----------------------------------------------------------------------------------------------------------------------------------------------------
    Top down recursion from questionnaire template to the lowest sections, each time noting which section ObjDef has which child section ObjDefs.
    For each combination the ARCM questionnaire template section structure is compared. If there is a matching pair of section AppObjs the the
    GUID of the ARCM child section is used for the export info, otherwise a new GUID is calculated.
    The original secion ObjDef GUID is never used.
 -----------------------------------------------------------------------------------------------------------------------------------------------------*/
function createSectionAndQuestionExportInfoRecursive(oSuperiorSectionDef, aSections, aAllImplicitSections, aAllImplicitQuestions, aARCMSections, aARCMQuestions, aSectionExportInfoResult, aQuestionExportInfoResult) {
    
    //---- search matches between ARIS ObjDefs and existing ARCM AppObjs
    var hm_section_objDef2ArcmObjectHashMap = new java.util.HashMap();
    for (var h=0; h<aSections.length; h++) {
        for (var i=0; i<aARCMSections.length; i++) {
            var sArisGUID = aARCMSections[i].getValueAttribute("objdef_guid").getRawValue();
            if (aSections[h].GUID() == sArisGUID) {
                hm_section_objDef2ArcmObjectHashMap.put(aSections[h], aARCMSections[i]);
            }
        }
    }
  
    //---- recursion for explicit sections
    for (var j=0; j<aSections.length; j++) {
        
        var aARISSubSections = new Array();
        aARISSubSections = aARISSubSections.concat(g_section2subsectionsHashMap.get(aSections[j]));
        
        var aARCMSubSections = new Array();
        var aARCMQuestions = new Array();
        var matchingARCMObject = hm_section_objDef2ArcmObjectHashMap.get(aSections[j]);
        if (matchingARCMObject != null) {
            aARCMSubSections = convertHashSetToJSArray( matchingARCMObject.getListAttribute("subSections").getConnectedObjects() );
            aARCMQuestions = convertHashSetToJSArray( matchingARCMObject.getListAttribute("questions").getConnectedObjects() );
        }
        
        //recursion
        createSectionAndQuestionExportInfoRecursive(aSections[j], aARISSubSections, aAllImplicitSections, aAllImplicitQuestions, aARCMSubSections, aARCMQuestions, aSectionExportInfoResult, aQuestionExportInfoResult);
        
        //section export info
        var arcmGUID = "";
        var matchingARCMObject = hm_section_objDef2ArcmObjectHashMap.get(aSections[j]);
        if (matchingARCMObject != null) {
            arcmGUID = matchingARCMObject.getValueAttribute("guid").getRawValue();
        } else {
            arcmGUID = createGuid();
        }
             
        var expInf = new exportInfo(aSections[j], oSuperiorSectionDef, null, arcmGUID);
        aSectionExportInfoResult.push(expInf);
    }
    
    //---- determine the implicit section placings which are relevant for the current superior section (if null then we look at export info for top level sections)
    var aImplicitSections = new Array();
    var aImplicitSectionsActivatingQuestions = new Array();
    for (var i=0; i<aAllImplicitSections.length; i++) {
        var relevantImplicitSectionObjDef = null;
        if (aAllImplicitSections[i].oPositioningSection == null && oSuperiorSectionDef == null) {
            relevantImplicitSectionObjDef = aAllImplicitSections[i].oObjDefToPlace;
            aImplicitSections.push(relevantImplicitSectionObjDef);
            aImplicitSectionsActivatingQuestions.push(aAllImplicitSections[i].oActivatingQuestion);
        } else if ( aAllImplicitSections[i].oPositioningSection != null
                    && oSuperiorSectionDef != null
                    && aAllImplicitSections[i].oPositioningSection.GUID() == oSuperiorSectionDef.GUID() ) {
            relevantImplicitSectionObjDef = aAllImplicitSections[i].oObjDefToPlace;
            aImplicitSections.push(relevantImplicitSectionObjDef);
            aImplicitSectionsActivatingQuestions.push(aAllImplicitSections[i].oActivatingQuestion);     
        } 
        //update mapping from section ObjDef to ARCMObjectDefinition
        if (relevantImplicitSectionObjDef != null) {
            for (var q=0; q<aARCMSections.length; q++) {
                var sArisGUID = aARCMSections[q].getValueAttribute("objdef_guid").getRawValue();
                if (relevantImplicitSectionObjDef.GUID() == sArisGUID) {
                    hm_section_objDef2ArcmObjectHashMap.put(relevantImplicitSectionObjDef, aARCMSections[q]);
                }
            }
        }
    }
    
    //---- recursion for implicit sections
    for (var k=0; k<aImplicitSections.length; k++) {
        
        var aARISSubSections = new Array();
        aARISSubSections = aARISSubSections.concat(g_section2subsectionsHashMap.get(aImplicitSections[k]));  
        var matchingARCMObject = hm_section_objDef2ArcmObjectHashMap.get(aImplicitSections[k]);
        
        var aARCMSubSections = new Array();
        var aARCMQuestions = new Array();
        if (matchingARCMObject != null) {
            aARCMSubSections = convertHashSetToJSArray( matchingARCMObject.getListAttribute("subSections").getConnectedObjects() );
            aARCMQuestions = convertHashSetToJSArray( matchingARCMObject.getListAttribute("questions").getConnectedObjects() );
        }
        
        //recursion
        createSectionAndQuestionExportInfoRecursive(aImplicitSections[k], aARISSubSections, aAllImplicitSections, aAllImplicitQuestions, aARCMSubSections, aARCMQuestions, aSectionExportInfoResult, aQuestionExportInfoResult);
        
        //section export info
        var arcmGUID = "";
        var matchingARCMObject = hm_section_objDef2ArcmObjectHashMap.get(aImplicitSections[k]);
        if (matchingARCMObject != null) {
            arcmGUID = matchingARCMObject.getValueAttribute("guid").getRawValue();
        } else {
            arcmGUID = createGuid();
        }
        
        //make sure that this implicit section is only activated by this one (explicit) activating question only (and not by other questions too which share the same option...)
        //find a matching exportInfo - if there is nonen yet then create a new one
        var expInf = null;
        for (var p=0; p<aSectionExportInfoResult.length; p++) {    
            var implicitSectionMatch = aImplicitSections[k].GUID() == aSectionExportInfoResult[p].exportObjDef.GUID();
            var superiorSectionMatch;
            if (oSuperiorSectionDef !== null && aSectionExportInfoResult[p].superiorSectionObjDef != null) {
                superiorSectionMatch = oSuperiorSectionDef.GUID() == aSectionExportInfoResult[p].superiorSectionObjDef.GUID()
            } else {
                superiorSectionMatch = oSuperiorSectionDef == null && aSectionExportInfoResult[p].superiorSectionObjDef == null;
            }

            if (implicitSectionMatch && superiorSectionMatch) {
                expInf = aSectionExportInfoResult[p];
            }
        }
        
        if (expInf != null) {
            var aActivatingQuestionsArray = expInf.aActivatingQuestionObjDefs;
            aActivatingQuestionsArray.push(aImplicitSectionsActivatingQuestions[k]);
        }
        else {
            var aActivatingQuestionsArray = new Array();
            aActivatingQuestionsArray.push(aImplicitSectionsActivatingQuestions[k]);
            expInf = new exportInfo(aImplicitSections[k], oSuperiorSectionDef, aActivatingQuestionsArray, arcmGUID);
            aSectionExportInfoResult.push(expInf);
        }   
    }
    
    //---- determine the implicit question placings which are relevant for the current superior section
    var aImplicitQuestions = new Array();
    var aImplicitQuestionsActivatingQuestions = new Array();
    for (var m=0; m<aAllImplicitQuestions.length; m++) {
        if (oSuperiorSectionDef != null && aAllImplicitQuestions[m].oPositioningSection.GUID() == oSuperiorSectionDef.GUID()) {
            aImplicitQuestions.push(aAllImplicitQuestions[m].oObjDefToPlace);
            aImplicitQuestionsActivatingQuestions.push(aAllImplicitQuestions[m].oActivatingQuestion);
        }
    }
    
    //---- question export info for explicit questions
    var aExplicitQuestionObjDefs = g_section2questionsHashMap.get(oSuperiorSectionDef);
    if (aExplicitQuestionObjDefs != null) {
        for (var o=0; o<aExplicitQuestionObjDefs.length; o++) {
            var expInf = new exportInfo(aExplicitQuestionObjDefs[o], oSuperiorSectionDef, null, aExplicitQuestionObjDefs[o].GUID());
            aQuestionExportInfoResult.push(expInf);
        }
    }
    
    //---- question export info for implicit questions
    for (var n=0; n<aImplicitQuestions.length; n++) {
        var arcmGUID = "";
        for (var o=0; o<aARCMQuestions.length; o++) {
            if (aImplicitQuestions[n].GUID() == aARCMQuestions[o].getValueAttribute("objdef_guid").getRawValue()) {
                //matching question found
                arcmGUID = aARCMQuestions[o].getValueAttribute("guid").getRawValue()
            }
        }
        //no matching question found
        if (arcmGUID == "") {
            arcmGUID = createGuid();
        }
        
        //make sure that this implicit question is only activated by this one (explicit) activating question only (and not by other questions too which share the same option...)
        //find a matching exportInfo - if there is nonen yet then create a new one
        var expInf = null;
        for (var p=0; p<aQuestionExportInfoResult.length; p++) {     
            var implicitQuestionMatch = aImplicitQuestions[n].GUID() == aQuestionExportInfoResult[p].exportObjDef.GUID();
            var superiorSectionMatch;
            if (oSuperiorSectionDef !== null && aQuestionExportInfoResult[p].superiorSectionObjDef != null) {
                superiorSectionMatch = oSuperiorSectionDef.GUID() == aQuestionExportInfoResult[p].superiorSectionObjDef.GUID()
            } else {
                superiorSectionMatch = oSuperiorSectionDef == null && aQuestionExportInfoResult[p].superiorSectionObjDef == null;
            }

            if (implicitQuestionMatch && superiorSectionMatch) {
                expInf = aQuestionExportInfoResult[p];
            }
        }
        
        if (expInf != null) {
            var aActivatingQuestionsArray = expInf.aActivatingQuestionObjDefs;
            aActivatingQuestionsArray.push(aImplicitQuestionsActivatingQuestions[n]);
        }
        else {
            var aActivatingQuestionsArray = new Array();
            aActivatingQuestionsArray.push(aImplicitQuestionsActivatingQuestions[n]);
            expInf = new exportInfo(aImplicitQuestions[n], oSuperiorSectionDef, aActivatingQuestionsArray, arcmGUID);
            aQuestionExportInfoResult.push(expInf);
        }
    }
}



/*----------------------------------------------------------------------------------------------------------------------------------------------------
    Update the GUID mapping for all documents of the given master objects (Risk, Questio, Control execution task etc.).
    If there is no mapping for a document matching name and link and linked to the given master object (i.e. it is imported the first time) 
    then create a new ARCM GUID for it and store this as mapping
 -----------------------------------------------------------------------------------------------------------------------------------------------------*/
function updateDocumentCache(sMasterGUID, sListAttributeID, sDocName, sDocLink) {
    
    //stand-alone execution of export report
    if (ARCM == null || g_arcmEnvironmentID == null) {     
        var arcmGUID = Packages.java.util.UUID.randomUUID().toString();
        storeDocumentGUIDInGUIDMappingCache(sDocName, sDocLink, sMasterGUID, arcmGUID);
        return;
    }
    
    //cache GUID->ARCMAppObject
    var masterAppObject = getFromObjDefMappingCache(sMasterGUID);
    if (masterAppObject == null) {
        
        var oGUIDRequestResultList = ARCM.findByGUID(sMasterGUID, g_arcmEnvironmentID, g_nLoc);
        if (!oGUIDRequestResultList.isEmpty()) {
            masterAppObject = oGUIDRequestResultList.get(0);
        } else {
            var arcmGUID = createGuid();
            storeDocumentGUIDInGUIDMappingCache(sDocName, sDocLink, sMasterGUID, arcmGUID);
            return;
        }
        storeInObjDefMappingCache(sMasterGUID, masterAppObject);
        
        //find matching document AppObject at master AppObject
        var aDocumentAppObjects = convertHashSetToJSArray( masterAppObject.getListAttribute(sListAttributeID).getConnectedObjects() );
        for (var i=0; i<aDocumentAppObjects.length; i++) {
            storeInMasterObjectCache(masterAppObject, aDocumentAppObjects[i]);
        }
    }
    
    //look up ARCM GUID for section from section AppObject cache
    var aDocuments = getFromMasterObjectCache(masterAppObject);
    for (var i=0; i<aDocuments.length && arcmGUID == null; i++) {   
        var sARCMDocName = aDocuments[i].getValueAttribute("name").getRawValue();
        var sARCMDocLink = aDocuments[i].getValueAttribute("link").getRawValue();
        if ( ((sARCMDocName == null) && (sDocName == null)) || (sARCMDocName != null && sARCMDocName.equals(sDocName)) 
             && ((sARCMDocLink == null) && (sDocLink == null)) || (sARCMDocLink != null && sARCMDocLink.equals(sDocLink))) {
            arcmGUID = aDocuments[i].getValueAttribute("guid").getRawValue(); 
        }
    }
    
    //no ARCM GUID for document found in ARCM objects -> store a new GUID for the document
    if (arcmGUID == null) {
        arcmGUID = createGuid();
    }
    
    storeDocumentGUIDInGUIDMappingCache(sDocName, sDocLink, sMasterGUID, arcmGUID);
}

/*----------------------------------------------------------------------------------------------------------------------------------------------------
    Update the GUID mapping for all documents of the given master objects (Risk, Questio, Control execution task etc.).
    If there is no mapping for a document matching name and link and linked to the given master object (i.e. it is imported the first time) 
    then create a new ARCM GUID for it and store this as mapping
 -----------------------------------------------------------------------------------------------------------------------------------------------------*/
function updatePolicyReviewTaskCache(sPolicyDefinitionGUID) {
    
    //stand-alone execution of export report
    if (ARCM == null || g_arcmEnvironmentID == null) {     
        var arcmGUID = Packages.java.util.UUID.randomUUID().toString();
        storePolicyReviewTaskGUIDInGUIDMappingCache(sPolicyDefinitionGUID, arcmGUID);
        return;
    }
    
	//load policy definition from ARCM - no caching required since this should only be donce once per PD
    //normally ARCM.findByGUID() should do but for older migrated databases we have the problem that 
    //a policydefinition and its policyreviewtask can have the same GUID
    var restrictionFactory = ARCM.getQueryRestrictionFactory();
    var query = ARCM.createQuery("POLICYDEFINITION", g_nLoc);
    query.addRestriction(restrictionFactory.eq("guid", sPolicyDefinitionGUID));
    var clientSignList = new java.util.ArrayList();
    clientSignList.add(g_arcmEnvironmentID);
    query.addClientRestriction (clientSignList);
    
    var oGUIDRequestResultList = query.getResult();
    if (!oGUIDRequestResultList.isEmpty()) {
        policyDefinitionAppObject = oGUIDRequestResultList.get(0);
    } else {
        var arcmGUID = createGuid(); 
        storePolicyReviewTaskGUIDInGUIDMappingCache(sPolicyDefinitionGUID, arcmGUID);
        return;
    }
    
    //load policy review task AppObject at policy definition AppObject
    var aPolicyReviewTaskAppObjects = convertHashSetToJSArray( policyDefinitionAppObject.getListAttribute("policyreviewtask").getConnectedObjects() );
    var aPolicyReviewTaskAppObject = aPolicyReviewTaskAppObjects[0];
    arcmGUID = aPolicyReviewTaskAppObject.getValueAttribute("guid").getRawValue(); 
    
    storePolicyReviewTaskGUIDInGUIDMappingCache(sPolicyDefinitionGUID, arcmGUID);
}