/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

var OUTPUTFILENAME = Context.getSelectedFile();

var g_nLoc = Context.getSelectedLanguage();

// needed for uniqueness check of controls and is initialized by function getObject(..)
var gRelevantModelsSet = new java.util.HashSet();
               
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
    var aControls = convertHashSetToJSArray(g_controlSet);
    
    
    //3. checks control
    var szControlOutput = "";
    for (var i=0; i<aControls.length; i++){
        
        var szInnerMsg = "";     
        //Test management checks
        if (g_bTestManagementExport) {
			
            
            szInnerMsg = addSingleValidationOutput( validateControlUniquenessOverBCDs(aControls[i]), szInnerMsg );
            szInnerMsg = addSingleValidationOutput( validateControlMandatoryAttributes(aControls[i]), szInnerMsg );
            szInnerMsg = addSingleValidationOutput( validateConnectedControlManagerGroupsCount(aControls[i]), szInnerMsg );
			if (g_approach == "RBA") {
				szInnerMsg = addSingleValidationOutput( validateConnectedRiskCount(aControls[i]), szInnerMsg );
			}
			if (g_approach == "CBA") {
				szInnerMsg = addSingleValidationOutput( validateConnectedFunctionsCount(aControls[i]), szInnerMsg );
				szInnerMsg = addSingleValidationOutput( validateConnectedRegulationCount(aControls[i]), szInnerMsg );
			}
        }
        //Risk management checks
        if (g_bRiskManagementExport) {
            //No control specific checks for risk management so far
        }
        
        //in case of error add the control object info
        if (!szInnerMsg.equals("")){            
            var szControlInfo = new java.lang.String(getString("TEXT_2")).replaceFirst("%0", aControls[i].Name(g_nLoc));
            var sCompleteSingleControlOutput = addObjectValidationInfo(szControlInfo, szInnerMsg, aControls[i], SPC1);
            szControlOutput = addCompleteObjectValidationOutput(szControlOutput, sCompleteSingleControlOutput); 
        }       
    }

    if (szControlOutput.equals("")) {szControlOutput = NO_ERROR_FOUND;}
    szControlOutput = addOutputHeader(szControlOutput, getString("TEXT_1"));
    
    
    //4. checks control execution tasks
    var aControlExecutionTask = convertHashSetToJSArray(g_controlExecutionTaskSet);
    var szControlTaskOutput = "";
    for (var i=0; i<aControlExecutionTask.length; i++){
        
        var szInnerMsg = "";     
        
        //Control management checks
        szInnerMsg = addSingleValidationOutput( validateControlExecutionTaskUniquenessOverBCDs(aControlExecutionTask[i]), szInnerMsg );
        szInnerMsg = addSingleValidationOutput( validateControlExecutionTaskMandatoryAttributes(aControlExecutionTask[i]), szInnerMsg );
        szInnerMsg = addSingleValidationOutput( validateControlExecutionFrequency(aControlExecutionTask[i]), szInnerMsg );
        szInnerMsg = addSingleValidationOutput( validateControlExecutionTaskStartEndDate(aControlExecutionTask[i], SPC1), szInnerMsg );
        szInnerMsg = addSingleValidationOutput( validateConnectedControlCount(aControlExecutionTask[i]), szInnerMsg );                            
        szInnerMsg = addSingleValidationOutput( validateConnectedControlExecutionOwnerGroupsCount(aControlExecutionTask[i]), szInnerMsg );
        szInnerMsg = addSingleValidationOutput( validateConnectedOrgUnitCount(aControlExecutionTask[i]), szInnerMsg );

        //in case of error add the control execution task object info
        if (!szInnerMsg.equals("")){            
            var szControlInfo = new java.lang.String(getString("TEXT_9")).replaceFirst("%0", aControlExecutionTask[i].Name(g_nLoc));
            var sCompleteSingleControlOutput = addObjectValidationInfo(szControlInfo, szInnerMsg, aControlExecutionTask[i], SPC1);
            szControlTaskOutput = addCompleteObjectValidationOutput(szControlTaskOutput, sCompleteSingleControlOutput); 
        }                    
    }  
    
    if (szControlTaskOutput.equals("")) {szControlTaskOutput = NO_ERROR_FOUND;}
    szControlTaskOutput = addOutputHeader(szControlTaskOutput, getString("TEXT_8"));
    
    
    //5. output
    szOutput += szControlOutput;
	//if there are errors for both control and control execution task then make sure they are properly separated
	if (!szControlOutput.equals("") && !szControlTaskOutput.equals("")) {szOutput += "\r\n" + "\r\n";}
    szOutput += szControlTaskOutput;
	
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
//-------------------------- Test management validates --------------------------------
function validateControlMandatoryAttributes(pControl){
    var szMsg = "";
    var aMandatories = getMandatoryFieldsForControl();   
    return validateMandatoryObjectAttributes(pControl, aMandatories, SPC1, getString("TEXT_4")); 
}


function validateControlUniquenessOverBCDs(pControl) {
    
    //determine all BCDs of the control
    var aControlObjOccs = pControl.OccList();
    for (var i=0; i<aControlObjOccs.length; i++) {
        if ( aControlObjOccs[i].Model().TypeNum() == Constants.MT_BUSY_CONTR_DGM ) {
            gRelevantModelsSet.add( aControlObjOccs[i].Model() );
        }
    }
    
    var szMsg = validateObjectUniquenessOverModels( pControl, getString("TEXT_3"), convertHashSetToJSArray(gRelevantModelsSet), SPC1 );
    return szMsg;  
}


function validateConnectedRiskCount(pControl){
    
    //get all linked export relevant risks
    //var aLinkedExportRelevantRisks = hm_linkedExportRelevantRisks.get(pControl);
    var aLinkedExportRelevantRisks = g_control2risksHashMap.get(pControl);
    //check cardinality
	if (g_approach == "RBA") {
		return getConnectionCountValidationOutput(OCCURRENCE_MIN_CONTROL_TO_RISKS_RBA, OCCURRENCE_MAX_CONTROL_TO_RISKS_RBA, pControl, aLinkedExportRelevantRisks, getString("TEXT_5"), SPC1);
	}
	if (g_approach == "CBA") {
		return getConnectionCountValidationOutput(OCCURRENCE_MIN_CONTROL_TO_RISKS_CBA, OCCURRENCE_MAX_CONTROL_TO_RISKS_CBA, pControl, aLinkedExportRelevantRisks, getString("TEXT_5"), SPC1);
	}
	return "";
   
}

function validateConnectedControlManagerGroupsCount(pControl){
    
   //var aConnectedGroups = getConnectedObjects(pControl,Constants.EDGES_IN, Constants.CT_IS_TECH_RESP_3, Constants.OT_PERS_TYPE); 
   var aConnectedGroups = g_control2controlManagerHashMap.get(pControl);
   return getConnectionCountValidationOutput(OCCURRENCE_MIN_CONTROL_TO_CONTROLMANAGER, OCCURRENCE_MAX_CONTROL_TO_CONTROLMANAGER, pControl, aConnectedGroups, getString("TEXT_7"), SPC1);     
     
}

function validateConnectedFunctionsCount(pControl){

   //var aConnectedFunctions = getConnectedObjects(pControl,Constants.EDGES_OUT, Constants.CT_OCCUR, Constants.OT_FUNC); 
   var aConnectedFunctions = new Array();
   if (g_control2processesHashMap.get(pControl) != null) {aConnectedFunctions = g_control2processesHashMap.get(pControl);}
   return  getConnectionCountValidationOutput(OCCURRENCE_MIN_CONTROL_TO_FUNCTION, OCCURRENCE_MAX_CONTROL_TO_FUNCTION, pControl, aConnectedFunctions, getString("TEXT_16"), SPC1);  
}

function validateConnectedRegulationCount(pControl){
    
    var aConnectedRegulations = g_control2regulationsHashMap.get(pControl);
    return  getConnectionCountValidationOutput(OCCURRENCE_MIN_CONTROL_TO_REGULATION, OCCURRENCE_MAX_CONTROL_TO_REGULATION, pControl, aConnectedRegulations, getString("TEXT_17"), SPC1);         
    
}

//-------------------------------------------------------------------------------------
//-------------------------- Control management validates --------------------------------
function validateControlExecutionTaskMandatoryAttributes(pControlExecutionTask){
    var szMsg = "";
    var aMandatories = getMandatoryFieldsForControlExecutionTask();   
    return validateMandatoryObjectAttributes(pControlExecutionTask, aMandatories, SPC1, getString("TEXT_11"));
}

function validateConnectedControlCount(pControlExecutionTask){
    var aLinkedControls = g_controlExecutionTask2controlHashMap.get(pControlExecutionTask);
    return getConnectionCountValidationOutput(OCCURRENCE_MIN_CONTROLEXECUTIONTASK_TO_CONTROL, OCCURRENCE_MAX_CONTROLEXECUTIONTASK_TO_CONTROL, pControlExecutionTask, aLinkedControls, getString("TEXT_12"), SPC1);     
}

function validateConnectedControlExecutionOwnerGroupsCount(pControlExecutionTask){
   var aConnectedGroups = g_controlExecutionTask2controlExecutionOwnerHashMap.get(pControlExecutionTask);
   return getConnectionCountValidationOutput(OCCURRENCE_MIN_CONTROLEXECUTIONTASK_TO_CONTROLEXECUTIONOWNER, OCCURRENCE_MAX_CONTROLEXECUTIONTASK_TO_CONTROLEXECUTIONOWNER, pControlExecutionTask, aConnectedGroups, getString("TEXT_13"), SPC1);       
}

function validateConnectedOrgUnitCount(pControlExecutionTask){
   var aConnectedOrgUnits = g_controlExecutionTask2orgunitHashMap.get(pControlExecutionTask);
   return getConnectionCountValidationOutput(OCCURRENCE_MIN_CONTROLEXECUTIONTASK_TO_ORGUNIT, OCCURRENCE_MAX_CONTROLEXECUTIONTASK_TO_ORGUNIT, pControlExecutionTask, aConnectedOrgUnits, getString("TEXT_14"), SPC1);       
}

function validateControlExecutionTaskStartEndDate(pControlExecutionTask, p_sSPC) {
	return validateStartEndDate(pControlExecutionTask, Constants.AT_CTRL_EXECUTION_TASK_START_DATE, Constants.AT_CTRL_EXECUTION_TASK_END_DATE, p_sSPC);
}

function validateControlExecutionFrequency(pControlExecutionTask) {
    var szMsg = "";
    if ( (isboolattributefalse(pControlExecutionTask, Constants.AT_EVENT_DRIVEN_CTRL_EXECUTION_ALLOWED, g_nLoc) 
         && pControlExecutionTask.Attribute(Constants.AT_CTRL_EXECUTION_TASK_FREQUENCY, g_nLoc).MeasureUnitTypeNum() == Constants.AVT_EVENT_DRIVEN)
         ) { 
        szMsg = SPC1 + getString("TEXT_15");
    }
    return szMsg;
}

function validateControlExecutionTaskUniquenessOverBCDs(pControlExecutionTask) {
    
    //determine all BCDs of the control execution task
    var aControlExecutionTaskObjOccs = pControlExecutionTask.OccList();
    for (var i=0; i<aControlExecutionTaskObjOccs.length; i++) {
        if ( aControlExecutionTaskObjOccs[i].Model().TypeNum() == Constants.MT_BUSY_CONTR_DGM ) {
            gRelevantModelsSet.add( aControlExecutionTaskObjOccs[i].Model() );
        }
    }
    
    var szMsg = validateObjectUniquenessOverModels( pControlExecutionTask, getString("TEXT_10"), convertHashSetToJSArray(gRelevantModelsSet), SPC1 );
    return szMsg;  
}