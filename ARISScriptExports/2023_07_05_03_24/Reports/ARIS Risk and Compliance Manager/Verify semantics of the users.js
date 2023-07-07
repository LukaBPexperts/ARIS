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
    var aUser = convertHashSetToJSArray(g_userSet);
    
    
    //3. checks
    for (var i=0; i<aUser.length; i++){
        
        var szInnerMsg = "";

        szInnerMsg = addSingleValidationOutput( validateUserMandatoryAttributes(aUser[i]), szInnerMsg );
        szInnerMsg = addSingleValidationOutput( validateConnectedGroupsCount(aUser[i]), szInnerMsg );

        if (!szInnerMsg.equals("")){
            
            var szUserInfo = new java.lang.String(getString("TEXT_2")).replaceFirst("%0", aUser[i].Name(g_nLoc));
            var sCompleteSingleUserOutput = addObjectValidationInfo(szUserInfo, szInnerMsg, aUser[i], SPC1);
            szOutput = addCompleteObjectValidationOutput(szOutput, sCompleteSingleUserOutput);  
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
function validateUserMandatoryAttributes(pUser){
    var szMsg = "";
    var aMandatories = getMandatoryFieldsForUser();
    
    //check mandatory attributes
    szMsg = validateMandatoryObjectAttributes(pUser, aMandatories, SPC1, getString("TEXT_3"));

    //check the user name
    if(!validateUserName(pUser)) {
        szMsg = szMsg + "\r\n\n" + SPC1 + getString("TEXT_5");    
    }
    
    return szMsg; 
}


function validateConnectedGroupsCount(pUser){
    
   var aConnectedGroups = getConnectedObjects(pUser, Constants.EDGES_OUT, Constants.CT_EXEC_5, Constants.OT_PERS_TYPE); 
   return getConnectionCountValidationOutput(OCCURRENCE_MIN_USER_TO_GROUPS, OCCURRENCE_MAX_USER_TO_GROUPS, pUser, aConnectedGroups, getString("TEXT_4"), SPC1);       
    
}

/*
 * validates letter figure combination
 */
function validateUserName(pUser) {
    strName = pUser.Name(g_nLoc);
    result = strName.match("[^\u0021-\u002c\u002f\u003A-\u0040\u005B-\u005E\^Â°Â§Â´`Â²Â³{}\\~Âµ|?\t\n\r]*" );
    if (strName.length() == result.join().length ){
       return true; 
    }
    return false;

}