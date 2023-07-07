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

var OUTPUTFILENAME = Context.getSelectedFile();

var g_nLoc = Context.getSelectedLanguage();             
             
try {           
    main();
}
catch(ex) {
    setExceptionProperty(ex);
}

function main(){
    var szOutput = "";
    
    //1. approach and mode
	if(determineApproach() == false) {
        return;
    }
    if (!g_bTestManagementExport && !g_bRiskManagementExport) {
        determineModules();    
    }
    //if no mode selected then nothing to do
    if (!g_bTestManagementExport && !g_bRiskManagementExport) {return;}
    
    
    //2. selection
    if (g_bTestManagementExport && !g_testManagementHashSetsFilled) {
        getTestManagementSelectionByContext();
    }
    if (g_bRiskManagementExport && !g_riskManagementHashSetsFilled) {
        getRiskManagementSelectionByContext();
    }
    var aUserGroups = convertHashSetToJSArray(g_userGroupSet);
    
    
    //3. checks
    for (var i=0; i<aUserGroups.length; i++){
        
        var szInnerMsg = "";
        
        szInnerMsg = addSingleValidationOutput( validateUserGroupMandatoryAttributes(aUserGroups[i]), szInnerMsg );
        szInnerMsg = addSingleValidationOutput( validateConnectedRolesCount(aUserGroups[i]), szInnerMsg );
		szInnerMsg = addSingleValidationOutput( validateConnectedUsersCount(aUserGroups[i]), szInnerMsg );
               
        if (!szInnerMsg.equals("")){
            
            var szUserGroupInfo = new java.lang.String(getString("TEXT_2")).replaceFirst("%0", aUserGroups[i].Name(g_nLoc));
            var sCompleteSingleUserGroupOutput = addObjectValidationInfo(szUserGroupInfo, szInnerMsg, aUserGroups[i], SPC1);
            szOutput = addCompleteObjectValidationOutput(szOutput, sCompleteSingleUserGroupOutput);     
        }
    }   
    
    if (szOutput.equals("")) {szOutput = NO_ERROR_FOUND;}
    szOutput = addOutputHeader(szOutput, getString("TEXT_1"));
    
    
    //4. output
    if (Context.getProperty("aam_semcheck") != null) {    
        // Evaluation by macro      
        if (Context.getProperty("reportdata") == null) {
            writeErrorreport(szOutput, OUTPUTFILENAME);
        } else {
            Context.setProperty("reportdata", szOutput);
        }
    } else {
        // Evaluation by report
        writeErrorreport(szOutput, OUTPUTFILENAME);
    }
}

//------------------------------------------------------------------------------------
//----------------------------- Validate functions -----------------------------------
function validateUserGroupMandatoryAttributes(pUserGroup){
    var szMsg = "";
    var aMandatories = getMandatoryFieldsForUserGroup();
    return validateMandatoryObjectAttributes(pUserGroup, aMandatories, SPC1, getString("TEXT_3")); 
}

function validateConnectedRolesCount(pUserGroup){
    
   var aConnectedGroups = getConnectedObjects(pUserGroup, Constants.EDGES_IN, Constants.CT_GENERAL_2, Constants.OT_PERS_TYPE); 
   return getConnectionCountValidationOutput(0, 1, pUserGroup, aConnectedGroups, getString("TEXT_4"), SPC1);       
    
}

function validateConnectedUsersCount(pUserGroup){
    
   var aConnectedUsers = getConnectedObjects(pUserGroup, Constants.EDGES_IN, Constants.CT_EXEC_5, Constants.OT_PERS); 
   return getConnectionCountValidationOutput(OCCURRENCE_MIN_USER_TO_GROUPS, OCCURRENCE_MAX_USER_TO_GROUPS, pUserGroup, aConnectedUsers, getString("TEXT_5"), SPC1);       
    
}