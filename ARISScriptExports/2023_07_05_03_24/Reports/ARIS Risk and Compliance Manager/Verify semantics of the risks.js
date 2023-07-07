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

// runtime improvement for attribute check
var g_aTestManagementAttributeConditions = new Array();
var g_aRiskManagementAttributeConditions = new Array();
var g_aReducedRiskManagementAttributeConditions = new Array();

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
    var aRisks = convertHashSetToJSArray(g_riskSet);
    
    
    //3. checks
    //determine the attribute conditions up front
    if (g_bTestManagementExport) {
        g_aTestManagementAttributeConditions = getMandatoryFieldsForRisk();
    }
    if (g_bRiskManagementExport) {
        g_aRiskManagementAttributeConditions = getMandatoryRiskManagementFieldsForRisk();
        if (g_bTestManagementExport) {
            //Work with the remaining conditions only
            g_aReducedRiskManagementAttributeConditions = getRemainingConditions(g_aRiskManagementAttributeConditions, g_aTestManagementAttributeConditions);
        }
    }
    
    //iterate on the risks
    for (var i=0; i<aRisks.length; i++) {
        
        var szInnerMsg = "";
        //Checks for both modes
        if (g_bTestManagementExport || g_bRiskManagementExport) {           
            szInnerMsg = addSingleValidationOutput( validateRiskUniquenessOverBCDs(aRisks[i]), szInnerMsg );
            szInnerMsg = addSingleValidationOutput( validateConnectedRiskManagerCount(aRisks[i]), szInnerMsg );
			if (g_approach == "RBA") {
				szInnerMsg = addSingleValidationOutput( validateConnectedFunctionsCount(aRisks[i]), szInnerMsg );
			}    
        }
        //Test management checks
        if (g_bTestManagementExport && g_processedRisksTestMode.contains(aRisks[i])) {
            szInnerMsg = addSingleValidationOutput( validateRiskMandatoryAttributes(aRisks[i]), szInnerMsg );
            szInnerMsg = addSingleValidationOutput( validateConnectedControlsCount(aRisks[i]), szInnerMsg );
        }
        //Risk management checks
        if (g_bRiskManagementExport && g_processedRisksRiskMode.contains(aRisks[i])) {
            //1. Risk management must be set to true -> described in Semantic_Properties
            if (g_bTestManagementExport && g_processedRisksTestMode.contains(aRisks[i])) {
                szInnerMsg = addSingleValidationOutput( validateReducedRiskManagementRiskMandatoryAttributes(aRisks[i]), szInnerMsg );
            }
            else {
                szInnerMsg = addSingleValidationOutput( validateRiskManagementRiskMandatoryAttributes(aRisks[i]), szInnerMsg );
            }
            //2. Risk management must be set to true -> cannot be described in Semantic_Properties -> hard-coded HACK
            if ( isboolattributetrue(aRisks[i], Constants.AT_GRC_RISK_MANAGEMENT_RELEVANT , g_nLoc) ) {
                szInnerMsg = addSingleValidationOutput( validateRiskStartEndDate(aRisks[i]), szInnerMsg );
                szInnerMsg = addSingleValidationOutput( validateConnectedRiskOwnerCount(aRisks[i]), szInnerMsg );
                szInnerMsg = addSingleValidationOutput( validateConnectedRiskReviewerCount(aRisks[i]), szInnerMsg );
				if (g_approach == "CBA") {
					szInnerMsg = addSingleValidationOutput( validateConnectedFunctionsCount(aRisks[i]), szInnerMsg );
				}
            }
        }
        
        //in case of error add the risk object info
        if (!szInnerMsg.equals("")){
            var szRiskInfo = new java.lang.String(getString("TEXT_2")).replaceFirst("%0", aRisks[i].Name(g_nLoc));
            var sCompleteSingleRiskOutput = addObjectValidationInfo(szRiskInfo, szInnerMsg, aRisks[i], SPC1)
            szOutput = addCompleteObjectValidationOutput(szOutput, sCompleteSingleRiskOutput);
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


//-------------------------------------------------------------------------------------
//-------------------------- Validates for both modes ---------------------------------
function validateRiskUniquenessOverBCDs(pRisk){
    
    var oRelevantModelsSet = new java.util.HashSet();
    
    //determine all BCDs of the risk
    var aRiskObjOccs = pRisk.OccList();
    for (var i=0; i<aRiskObjOccs.length; i++) {
        if ( aRiskObjOccs[i].Model().TypeNum() == Constants.MT_BUSY_CONTR_DGM ) {
            oRelevantModelsSet.add( aRiskObjOccs[i].Model() );
        }
    }
    
    var szMsg = validateObjectUniquenessOverModels( pRisk, getString("TEXT_4"), convertHashSetToJSArray(oRelevantModelsSet), SPC1 );
    return szMsg;
}

function validateConnectedRiskManagerCount(pRisk){
    
   //var aConnectedGroups = getConnectedObjects(pRisk,Constants.EDGES_IN, Constants.CT_IS_TECH_RESP_3, Constants.OT_PERS_TYPE); 
   var aConnectedGroups = new Array();
   if (g_risk2riskManagerHashMap.get(pRisk) != null) {aConnectedGroups = aConnectedGroups.concat(g_risk2riskManagerHashMap.get(pRisk));}
   return getConnectionCountValidationOutput(OCCURRENCE_MIN_RISK_TO_RISKMANAGER, OCCURRENCE_MAX_RISK_TO_RISKMANAGER, pRisk, aConnectedGroups, getString("TEXT_6"), SPC1);       
    
}

function validateConnectedFunctionsCount(pRisk){

   //var aConnectedFunctions = getConnectedObjects(pRisk,Constants.EDGES_OUT, Constants.CT_OCCUR, Constants.OT_FUNC); 
   var aConnectedFunctions = new Array();
   if (g_risk2processesHashMap.get(pRisk) != null) {aConnectedFunctions = g_risk2processesHashMap.get(pRisk);}
   return  getConnectionCountValidationOutput(OCCURRENCE_MIN_RISK_TO_FUNCTION, OCCURRENCE_MAX_RISK_TO_FUNCTION, pRisk, aConnectedFunctions, getString("TEXT_5"), SPC1);      
}

//-------------------------------------------------------------------------------------
//-------------------------- Test management validates --------------------------------
function validateRiskMandatoryAttributes(pRisk){   
    return validateMandatoryObjectAttributes(pRisk, g_aTestManagementAttributeConditions, SPC1, getString("TEXT_3")); 
}

function validateConnectedControlsCount(pRisk){

   var aConnectedControls = new Array();
   if (g_risk2controlsHashMap.get(pRisk) != null) {aConnectedControls = g_risk2controlsHashMap.get(pRisk);}
   return  getConnectionCountValidationOutput(OCCURRENCE_MIN_RISK_TO_CONTROL, OCCURRENCE_MAX_RISK_TO_CONTROL, pRisk, aConnectedControls, getString("TEXT_14"), SPC1);      
}

//-------------------------------------------------------------------------------------
//-------------------------- Risk management validates --------------------------------
function validateRiskManagementRiskMandatoryAttributes(pRisk){
    return validateMandatoryObjectAttributes(pRisk, g_aRiskManagementAttributeConditions, SPC1, getString("TEXT_3"));
}

function validateRiskStartEndDate(pRisk) {
	return validateStartEndDate(pRisk, Constants.AT_GRC_START_DATE_OF_RISK_ASSESSMENTS, Constants.AT_GRC_END_DATE_OF_RISK_ASSESSMENTS, SPC1);
}

function validateReducedRiskManagementRiskMandatoryAttributes(pRisk){
    return validateMandatoryObjectAttributes(pRisk, g_aReducedRiskManagementAttributeConditions, SPC1, getString("TEXT_3"));
}


function validateConnectedRiskOwnerCount(pRisk) {

    var aConnectedRiskOwners = g_risk2riskOwnerHashMap.get(pRisk);
    return  getConnectionCountValidationOutput(RM_OCCURRENCE_MIN_RISK_TO_RISKOWNER, RM_OCCURRENCE_MAX_RISK_TO_RISKOWNER, pRisk, aConnectedRiskOwners, getString("TEXT_8"), SPC1);
}


function validateConnectedRiskReviewerCount(pRisk) {

    var aConnectedRiskReviewers = g_risk2riskReviewerHashMap.get(pRisk);
    return  getConnectionCountValidationOutput(RM_OCCURRENCE_MIN_RISK_TO_RISKREVIEWER, RM_OCCURRENCE_MAX_RISK_TO_RISKREVIEWER, pRisk, aConnectedRiskReviewers, getString("TEXT_9"), SPC1);
}