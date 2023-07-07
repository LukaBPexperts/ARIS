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

var g_datePeriodCache      = new java.util.HashMap();     // global cache of calculated date periods - Format: ObjDef | DatePeriodContainer
    
try {           
    main();
}
catch(ex) {
    setExceptionProperty(ex);
}

function main() {
    var szOutput = "";
   
    //1. approach and mode
	if(determineApproach() == false) {
        return;
    }
    if (!g_bAuditManagementExport) {
        determineModules();    
    }
    //if no mode selected then nothing to do
    if (!g_bAuditManagementExport) {return;}
    
    
    //2. selection
    getAuditManagementSelectionByContext();
    var aAuditTemplates = convertHashSetToJSArray(g_auditTemplateSet);
    
    
    //3. checks
    for (var i=0; i<aAuditTemplates.length; i++) {
        
        var szInnerMsg = "";
    
        //---- audit template itself
        var sAuditTemplateErrors = validateAuditTemplate(aAuditTemplates[i], SPC1);
        if (sAuditTemplateErrors != null && sAuditTemplateErrors.length > 0) {
            if (szInnerMsg.length > 0) {szInnerMsg += "\r\n" + "\r\n";}
            szInnerMsg += sAuditTemplateErrors;
        }
        
        //---- all steps and substeps
        var oFoundStepsSet = new Packages.java.util.HashSet();
        //determine all steps beloning to the template directly or indirectly
        //and search for cycles
        var sOverallStepStructureErrors = validateStepStructure(aAuditTemplates[i], oFoundStepsSet, SPC1);
        if ( sOverallStepStructureErrors != null && sOverallStepStructureErrors.length > 0 ) {
            if (szInnerMsg.length > 0) {szInnerMsg += "\r\n" + "\r\n";}
            szInnerMsg += sOverallStepStructureErrors;
        }
        
        //now the individual step checks
        var sOverallStepParentErrors = "";
        for (var it = oFoundStepsSet.iterator(); it.hasNext();) {
            var oStep = it.next();
            
            //do a check on linked parents for each single step
            var sIndiviualCombinedStepErrors = "";
            
            var sStepParentErrors = validateStepParents(oStep, SPC1);
            if (sStepParentErrors != null && sStepParentErrors.length > 0) {
                if (sIndiviualCombinedStepErrors.length > 0) {sIndiviualCombinedStepErrors += "\r\n" + "\r\n";}
                sIndiviualCombinedStepErrors += sStepParentErrors;
                
                sOverallStepParentErrors += sStepParentErrors;
            }
            
            //do a check on linked subordinated objects and mandatory fields for each single step
            var sStepErrors = validateStepSubordinatesAndMandatoryFields(oStep, SPC1);
            if (sStepErrors != null && sStepErrors.length > 0) {
                if (sIndiviualCombinedStepErrors.length > 0) {sIndiviualCombinedStepErrors += "\r\n" + "\r\n";}
                sIndiviualCombinedStepErrors += sStepErrors;
            }
            
            
            //in case of errors at assigned parent, subordinates or mandatory fields:
            //add also header with the section name
            if ( (sIndiviualCombinedStepErrors != null && sIndiviualCombinedStepErrors.length > 0) ) {
                var sStepHeader = getString("TEXT_2") + " \"" + oStep.Name(g_nLoc) + "\":";
                sIndiviualCombinedStepErrors = SPC1 + sStepHeader + "\r\n"
                                                  + getUnderline(sStepHeader, SPC1) + "\r\n"
                                                  + sIndiviualCombinedStepErrors;
                
                if (szInnerMsg.length > 0) {szInnerMsg += "\r\n" + "\r\n"}
                szInnerMsg += sIndiviualCombinedStepErrors;
            }
            
        }
    
        //---- add header in case of errors
        if (!szInnerMsg.equals("")){
            var szAuditTemplateInfo = new java.lang.String(getString("TEXT_10")).replaceFirst("%", aAuditTemplates[i].Name(g_nLoc));
            var sCompleteSingleAuditTemplateOutput = addObjectValidationInfo(szAuditTemplateInfo, szInnerMsg, aAuditTemplates[i], SPC1);
            szOutput = addCompleteObjectValidationOutput(szOutput, sCompleteSingleAuditTemplateOutput);
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

//---------
function validateAuditTemplate(p_oAuditTemplate, p_sSPC) {

    var szMsg = "";
    
    //mandatory fields
    var sMandatoriesMsg = "";
    var aMandatories = getMandatoryFieldsForAuditTemplates();
    sMandatoriesMsg = validateMandatoryObjectAttributes(p_oAuditTemplate, aMandatories, p_sSPC);
        
    if (!sMandatoriesMsg.equals("")){
        szMsg = sMandatoriesMsg;
    }
    
    //duration attribute AT_MAX_TL_TIME must be filled with an value different from 0000:00:00:00
    for (var i=0; i<aMandatories.length; i++) {
        if (aMandatories[i] != "AT_MAX_TL_TIME") {continue;}
        var oAttribute = p_oAuditTemplate.Attribute(Constants.AT_MAX_TL_TIME , g_nLoc);
        var durationValue = oAttribute.getValue();
        if (durationValue == "0000:00:00:00") {
            szMsg = p_sSPC + oAttribute.Type() + " " + getString("COMMON_10");
        }
    }
    
    //preparation start date must not lie within audit period
    var aModels = getProjectScheduleModelWithObjDefOccurrence(p_oAuditTemplate);
    if (aModels != null && aModels.length > 0) {
        var auditDatePeriodContainer = calculateDatePeriod(aModels[0], p_oAuditTemplate, p_oAuditTemplate.OccListInModel(aModels[0])[0] );
        if (auditDatePeriodContainer != null) {
			var auditPreparationStartDate = p_oAuditTemplate.Attribute(Constants.AT_START_DATE_OF_AUDIT_PREPARATION, g_nLoc).MeasureValue();
			if ((auditPreparationStartDate != null) && (auditDatePeriodContainer.startDate != null)) {
				if ( auditDatePeriodContainer.startDate.compareTo(auditPreparationStartDate) < 0) {
					if (szMsg.length > 0) {szMsg += "\r\n" + "\r\n";}
						szMsg += p_sSPC + getString("TEXT_16");
				}
			}
		}
    }      
    
	//control period start date must lie before control prtiod end date
	var sStartEndDateMessage = validateStartEndDate(p_oAuditTemplate, Constants.AT_START_DATE_OF_CONTROL_PERIOD, Constants.AT_END_DATE_OF_CONTROL_PERIOD, SPC1);
	if (!sStartEndDateMessage.equals("")) {
        if (szMsg != "") {szMsg += "\r\n" + "\r\n";}
        szMsg += sStartEndDateMessage;
    }
	
    //linked subordinated objects
    var sLinkedTopStepsMsg = "";
    var aConnectedTopSteps = new Array();
    if (g_auditTemplate2AuditStepTemplateHashMap.get(p_oAuditTemplate) != null) {aConnectedTopSteps = g_auditTemplate2AuditStepTemplateHashMap.get(p_oAuditTemplate);}
    sLinkedTopStepsMsg = getConnectionCountValidationOutput(SM_OCCURRENCE_MIN_AUDIT_TEMPLATE_TO_TOP_STEP, SM_OCCURRENCE_MAX_AUDIT_TEMPLATE_TO_TOP_STEP, p_oAuditTemplate, aConnectedTopSteps, getString("TEXT_2"), p_sSPC);
    if (!sLinkedTopStepsMsg.equals("")) {
        if (szMsg.length > 0) {szMsg += "\r\n" + "\r\n";}
        szMsg += sLinkedTopStepsMsg;
    }
    
    var sLinkedOwnersMsg = "";
    var aConnectedOwners = new Array();
    if (g_auditTemplate2auditOwnerHashMap.get(p_oAuditTemplate) != null) {aConnectedOwners = g_auditTemplate2auditOwnerHashMap.get(p_oAuditTemplate);}
    sLinkedOwnersMsg = getConnectionCountValidationOutput(SM_OCCURRENCE_MIN_AUDIT_TEMPLATE_TO_AUDITOWNER, SM_OCCURRENCE_MAX_AUDIT_TEMPLATE_TO_AUDITOWNER, p_oAuditTemplate, aConnectedOwners, getString("TEXT_3"), p_sSPC);
    if (!sLinkedOwnersMsg.equals("")) {
        if (szMsg.length > 0) {szMsg += "\r\n" + "\r\n";}
        szMsg += sLinkedOwnersMsg;
    }
    
    var sLinkedReviewersMsg = "";
    var aConnectedReviewers = new Array();
    if (g_auditTemplate2auditReviewerHashMap.get(p_oAuditTemplate) != null) {aConnectedReviewers = g_auditTemplate2auditReviewerHashMap.get(p_oAuditTemplate);}
    sLinkedReviewersMsg = getConnectionCountValidationOutput(SM_OCCURRENCE_MIN_AUDIT_TEMPLATE_TO_AUDITREVIEWER, SM_OCCURRENCE_MAX_AUDIT_TEMPLATE_TO_AUDITREVIEWER, p_oAuditTemplate, aConnectedReviewers, getString("TEXT_4"), p_sSPC); 
    if (!sLinkedReviewersMsg.equals("")) {
        if (szMsg.length > 0) {szMsg += "\r\n" + "\r\n";}
        szMsg += sLinkedReviewersMsg;
    }
    
    var sLinkedAuditorsMsg = "";
    var aConnectedAuditors = new Array();
    if (g_auditTemplate2auditAuditorHashMap.get(p_oAuditTemplate) != null) {aConnectedAuditors = g_auditTemplate2auditAuditorHashMap.get(p_oAuditTemplate);}
    sLinkedAuditorsMsg = getConnectionCountValidationOutput(SM_OCCURRENCE_MIN_AUDIT_TEMPLATE_TO_AUDITAUDITOR, SM_OCCURRENCE_MAX_AUDIT_TEMPLATE_TO_AUDITAUDITOR, p_oAuditTemplate, aConnectedAuditors, getString("TEXT_5"), p_sSPC);
    if (!sLinkedAuditorsMsg.equals("")) {
        if (szMsg.length > 0) {szMsg += "\r\n" + "\r\n";}
        szMsg += sLinkedAuditorsMsg;
    }
    
    //linked scope
    var sLinkedScopesMsg = "";
    var aConnectedScopes = new Array();
    if (g_auditTemplate2regulationHashMap.get(p_oAuditTemplate) != null) {aConnectedScopes = aConnectedScopes.concat(g_auditTemplate2regulationHashMap.get(p_oAuditTemplate));}
    if (g_auditTemplate2orgUnitHashMap.get(p_oAuditTemplate) != null) {aConnectedScopes = aConnectedScopes.concat(g_auditTemplate2orgUnitHashMap.get(p_oAuditTemplate));}
    if (g_auditTemplate2processHashMap.get(p_oAuditTemplate) != null) {aConnectedScopes = aConnectedScopes.concat(g_auditTemplate2processHashMap.get(p_oAuditTemplate));}
    if (g_auditTemplate2appSysTypeHashMap.get(p_oAuditTemplate) != null) {aConnectedScopes = aConnectedScopes.concat(g_auditTemplate2appSysTypeHashMap.get(p_oAuditTemplate));}
    if (g_auditTemplate2riskCategoryHashMap.get(p_oAuditTemplate) != null) {aConnectedScopes = aConnectedScopes.concat(g_auditTemplate2riskCategoryHashMap.get(p_oAuditTemplate));}
    if (aConnectedScopes.length > 1) {
        if (szMsg.length > 0) {szMsg += "\r\n" + "\r\n";}
        szMsg += p_sSPC + getString("TEXT_15");
        for (var i=0; i<aConnectedScopes.length; i++) {
            szMsg += "\r\n" + p_sSPC + SPC1 + aConnectedScopes[i].Name(g_nLoc);
        }
    }
    
    return szMsg;
}

//---------
//checks for cycles in the section structure
function validateStepStructure(p_oAuditTemplate, p_oFoundStepsSet, p_sSPC) {

    var aAssignedSteps = g_auditTemplate2AuditStepTemplateHashMap.get(p_oAuditTemplate);
    
    if (aAssignedSteps != null) {
        
        var oFoundStepsSet = new Packages.java.util.HashSet();
        var oStepStack = new Packages.java.util.Stack();       
        var aStepCycleErrors = new Array();
        for (var i=0; i<aAssignedSteps.length; i++) {
            checkStepCyclesRecursively(aAssignedSteps[i], oFoundStepsSet, oStepStack, aStepCycleErrors);    
        }
        p_oFoundStepsSet.addAll(oFoundStepsSet);
    }
    
    //if errors were found
    var sCycleErrorString = "";
    for (var i=0; i<aStepCycleErrors.length; i++) {
        if (sCycleErrorString.length > 0) {sCycleErrorString += "\r\n";}
        sCycleErrorString += p_sSPC + aStepCycleErrors[i];
    }
    if (sCycleErrorString.length > 0) {
        sCycleErrorString = p_sSPC + getString("TEXT_6") + "\r\n" + sCycleErrorString;
    }
    
    return sCycleErrorString;    
}
//---------
function checkStepCyclesRecursively(p_oStep, p_oFoundStepsSet, p_oStepStack, p_aStepCycleErrors) {
    
    //if step is in the set of already handled steps then we have a cycle
    if (!p_oFoundStepsSet.add(p_oStep)) {
        var sCycleMessage = "";
        var bCycleBeginningReached = false;
        for (var it = p_oStepStack.iterator(); it.hasNext(); ) {
            
            var oStackItem = it.next();
            bCycleBeginningReached = bCycleBeginningReached || oStackItem.equals(p_oStep);
            if (!bCycleBeginningReached) {continue;}
            
            if (sCycleMessage.length > 0) {sCycleMessage += " - ";}
            sCycleMessage += oStackItem.Name(g_nLoc);
        }
        
        p_aStepCycleErrors.push(SPC1 + sCycleMessage);
        return;
    }
    
    //if not then put it on the stack...
    p_oStepStack.push(p_oStep);
    //...and check all children recursively...
    var aSubSteps = g_auditStepTemplate2subAuditStepTemplateHashMap.get(p_oStep);
    if (aSubSteps != null) {
        for (var i=0; i<aSubSteps.length; i++) {
            checkStepCyclesRecursively(aSubSteps[i], p_oFoundStepsSet, p_oStepStack, p_aStepCycleErrors);    
        }
    } 
    //...then remove from stack after recursion
    p_oStepStack.pop();    
}

//---------
function validateStepParents(p_oStep, p_sSPC) {
    
    var szMsg = "";
     
    //linked parents
    var aTemplates = g_auditStepTemplate2AuditTemplateHashMap.get(p_oStep);
    if (aTemplates != null && aTemplates.length > 1) {
        szMsg = p_sSPC + getString("TEXT_7");
        for (var i=0; i<aTemplates.length; i++) {
            szMsg += "\r\n" + p_sSPC + SPC1 + aTemplates[i].Name(g_nLoc);
        }
    }
    var aTopSteps = g_auditStepTemplate2topAuditStepTemplateHashMap.get(p_oStep);
    if (aTopSteps != null && aTopSteps.length > 1) {
        if (szMsg != "") {szMsg += "\r\n" + "\r\n";}
        szMsg = p_sSPC + getString("TEXT_8");
        for (var i=0; i<aTopSteps.length; i++) {
            szMsg += "\r\n" + p_sSPC + SPC1 + aTopSteps[i].Name(g_nLoc);
        }
        szMsg += "\r\n";
    }
    
    return szMsg;    
}

//---------
function validateStepSubordinatesAndMandatoryFields(p_oStep, p_sSPC) {
    
    var szMsg = "";
    
    //mandatory fields
    var sMandatoriesMsg = "";
    var aMandatories = getMandatoryFieldsForAuditStepTemplates();
    sMandatoriesMsg = validateMandatoryObjectAttributes(p_oStep, aMandatories, p_sSPC);
        
    if (!sMandatoriesMsg.equals("")){
        szMsg = sMandatoriesMsg;
    }
    
    //duration attributes AT_MAX_TL_TIME and AT_DES_PROC_TIME must be filled with an value different from 00:00:00
    for (var i=0; i<aMandatories.length; i++) {
        if (aMandatories[i] != "AT_MAX_TL_TIME" && aMandatories[i] != "AT_DES_PROC_TIME") {continue;}
        var durationValue;
        if (aMandatories[i] == "AT_MAX_TL_TIME") {  
            var oAttribute = p_oStep.Attribute(Constants.AT_MAX_TL_TIME , g_nLoc);
            var durationValue = oAttribute.getValue();
            if (durationValue == "0000:00:00:00") {
                szMsg = p_sSPC + oAttribute.Type() + " " + getString("COMMON_10");
            }
        }
        if (aMandatories[i] == "AT_DES_PROC_TIME") {
            var oAttribute = p_oStep.Attribute(Constants.AT_DES_PROC_TIME , g_nLoc);
            var durationValue = oAttribute.getValue();
            if (durationValue == "0000:00:00:00") {
                szMsg = p_sSPC + oAttribute.Type() + " " + getString("COMMON_10");
            }
        }
    }
    
    //linked subordinated objects
    var sLinkedStepOwnersMsg = "";
    var aConnectedStepOwners = new Array();
    if (g_auditStepTemplate2auditStepOwnerHashMap.get(p_oStep) != null) {aConnectedStepOwners = g_auditStepTemplate2auditStepOwnerHashMap.get(p_oStep);}
    sLinkedStepOwnersMsg = getConnectionCountValidationOutput(SM_OCCURRENCE_MIN_AUDIT_STEP_TEMPLATE_TO_AUDITSTEPOWNER, SM_OCCURRENCE_MAX_AUDIT_STEP_TEMPLATE_TO_AUDITSTEPOWNER, p_oStep, aConnectedStepOwners, getString("TEXT_9"), SPC1);      
    if (!sLinkedStepOwnersMsg.equals("")) {
        if (szMsg.length > 0) {szMsg += "\r\n" + "\r\n";}
        szMsg += sLinkedStepOwnersMsg;
    }
    
    //step may have only one single occurence inside a project schedule model
    var aModels = getProjectScheduleModelWithObjDefOccurrence(p_oStep);
    var szOccurrenceMsg = validateObjectUniquenessOverModels( p_oStep, getString("TEXT_2"), aModels, SPC1 );
    if (!szOccurrenceMsg.equals("")) {
        if (szMsg.length > 0) {szMsg += "\r\n" + "\r\n";}
        szMsg += sLinkedStepOwnersMsg;
    }

    //if there is only one occurrence then check if the "Start date"/"End date" period lies within the period of the superior step template or audit template
    if (aModels.length == 1) {      
        var szDatePeriodMsg = validateDatePeriods(aModels[0], p_oStep, SPC1);
        if (!szDatePeriodMsg.equals("")) { 
            if (szMsg.length > 0) {szMsg += "\r\n" + "\r\n";}
            szMsg += szDatePeriodMsg;
        }
    }
    
    //"Desired processing time" must not be greater than "Maximum total time"
    var datePeriodContainer = g_datePeriodCache.get(p_oStep); //cache was filled during execution of validateDatePeriods()
    if (p_oStep.Attribute( eval("Constants.AT_MAX_TL_TIME" ), g_nLoc).IsMaintained()
        && p_oStep.Attribute( eval("Constants.AT_DES_PROC_TIME" ), g_nLoc).IsMaintained() ) {
        
        var sDurationStringValue = p_oStep.Attribute( eval("Constants.AT_MAX_TL_TIME" ) , g_nLoc).GetValue(false);
        var lMaxTimeInMillis = getAttrValue_Duration(sDurationStringValue);
        
        var sDurationStringValue = p_oStep.Attribute( eval("Constants.AT_DES_PROC_TIME" ) , g_nLoc).GetValue(false);
        var lPlannedTimeInMillis = getAttrValue_Duration(sDurationStringValue);
        
        if (lPlannedTimeInMillis > lMaxTimeInMillis) {
            if (szMsg.length > 0) {szMsg += "\r\n" + "\r\n";}
            szMsg += p_sSPC + getString("TEXT_17");
        }
    }
    
    //linked scope
    var sLinkedScopesMsg = "";
    var aConnectedScopes = new Array();
    if (g_auditStepTemplate2regulationHashMap.get(p_oStep) != null) {aConnectedScopes = aConnectedScopes.concat(g_auditStepTemplate2regulationHashMap.get(p_oStep));}
    if (g_auditStepTemplate2orgUnitHashMap.get(p_oStep) != null) {aConnectedScopes = aConnectedScopes.concat(g_auditStepTemplate2orgUnitHashMap.get(p_oStep));}
    if (g_auditStepTemplate2processHashMap.get(p_oStep) != null) {aConnectedScopes = aConnectedScopes.concat(g_auditStepTemplate2processHashMap.get(p_oStep));}
    if (g_auditStepTemplate2appSysTypeHashMap.get(p_oStep) != null) {aConnectedScopes = aConnectedScopes.concat(g_auditStepTemplate2appSysTypeHashMap.get(p_oStep));}
    if (g_auditStepTemplate2riskCategoryHashMap.get(p_oStep) != null) {aConnectedScopes = aConnectedScopes.concat(g_auditStepTemplate2riskCategoryHashMap.get(p_oStep));}
    if (aConnectedScopes.length > 1) {
        if (szMsg.length > 0) {szMsg += "\r\n" + "\r\n";}
        szMsg += p_sSPC + getString("TEXT_15");
        for (var i=0; i<aConnectedScopes.length; i++) {
            szMsg += "\r\n" + p_sSPC + SPC1 + aConnectedScopes[i].Name(g_nLoc);
        }
    }
    
    return szMsg;    
}

function validateDatePeriods(p_oModel, p_oStep, p_sSPC) {
    
    var szMsg = "";
    
    var datePeriodContainer = calculateDatePeriod(p_oModel, p_oStep, p_oStep.OccListInModel(p_oModel)[0] );
    if (datePeriodContainer == null) {return szMsg;}
	if (datePeriodContainer.startDate == null) {return szMsg;}
    if (datePeriodContainer.endDate == null) {return szMsg;}
    
    //oSuperiorObjDef is always the audit template
    var oSuperiorObjDef = null;
    var oCurrentStep = p_oStep;
    var aSuperiorAuditTemplates = g_auditStepTemplate2AuditTemplateHashMap.get(oCurrentStep);
    while (aSuperiorAuditTemplates == null || aSuperiorAuditTemplates.length == 0) {
        var aSuperiorSteps = g_auditStepTemplate2topAuditStepTemplateHashMap.get(oCurrentStep);
        //there is no further step parent, abort loop
        if (aSuperiorSteps == null || aSuperiorSteps.length == 0) {
            break;
        }
        oCurrentStep = aSuperiorSteps[0];
        //abort complete validation function in case of cycle!
        if (oCurrentStep.equals(p_oStep)) {
            return "";
        }
        aSuperiorAuditTemplates = g_auditStepTemplate2AuditTemplateHashMap.get(oCurrentStep);
    }
    
    if (aSuperiorAuditTemplates == null || aSuperiorAuditTemplates.length == 0) {
        return szMsg;
    } else {
        oSuperiorObjDef = aSuperiorAuditTemplates[0];
    }
    
    /*
    //oSuperiorObjDef can be an audit template or an audit step template
    var oSuperiorObjDef = null;
    var aSuperiorAuditTemplates = g_auditStepTemplate2AuditTemplateHashMap.get(p_oStep);
    if (aSuperiorAuditTemplates != null && aSuperiorAuditTemplates.length == 1) {
        oSuperiorObjDef = aSuperiorAuditTemplates[0];
    } else {
        var aSuperiorSteps = g_auditStepTemplate2subAuditStepTemplateHashMap.get(p_oStep);
        if (aSuperiorSteps != null && aSuperiorSteps.length == 1) {
            oSuperiorObjDef = aSuperiorSteps[0];
        }
    }
    if (oSuperiorObjDef == null) {return szMsg;}
    */
    
    var aModels = getProjectScheduleModelWithObjDefOccurrence(oSuperiorObjDef);
    if (aModels == null || aModels.length != 1) {return szMsg;}
    superiorDatePeriodContainer = calculateDatePeriod(aModels[0], oSuperiorObjDef, oSuperiorObjDef.OccListInModel(aModels[0])[0] );
	if (superiorDatePeriodContainer == null) {return szMsg;}
	if (superiorDatePeriodContainer.startDate == null) {return szMsg;}
    if (superiorDatePeriodContainer.endDate == null) {return szMsg;}
   
	//check the date periods for p_oStep and its superior ObjDef
    if ( datePeriodContainer.startDate.compareTo(superiorDatePeriodContainer.startDate) < 0
         || datePeriodContainer.endDate.compareTo(superiorDatePeriodContainer.endDate) > 0) {

        var dateFormatter = new java.text.SimpleDateFormat(getString("TEXT_14"));
        
        if (szMsg.length > 0) {szMsg += "\r\n" + "\r\n";}
        szMsg += p_sSPC + getString("TEXT_11");
        szMsg += "\r\n" + p_sSPC + getString("TEXT_12") + " " + dateFormatter.format(datePeriodContainer.startDate) + " - " + dateFormatter.format(datePeriodContainer.endDate);
        szMsg += "\r\n" + p_sSPC + getString("TEXT_13") + " " + dateFormatter.format(superiorDatePeriodContainer.startDate) + " - " + dateFormatter.format(superiorDatePeriodContainer.endDate);
    }
    
    return szMsg;
}

//---------
function getProjectScheduleModelWithObjDefOccurrence(p_oObjDef) {
    var aObjOccs = p_oObjDef.OccList();
    var aModels = new Array();
    for (var i=0; i<aObjOccs.length; i++) {
        if ( aObjOccs[i].Model().TypeNum() == Constants.MT_PROJECT_SCHEDULE ) {
            aModels.push( aObjOccs[i].Model() );
        }
    }
    return aModels;
}

//---------
function DatePeriodContainer( startDate, endDate ) {    
    this.startDate = startDate;
    this.endDate = endDate;   
}

function calculateDatePeriod(oModel, oObjDef, oObjOcc) {
    
    if (oModel == null || oObjDef == null || oObjOcc == null) {
        return null;
    }
    
    //check if the calculation was already done for this oObjDef
    if (g_datePeriodCache.get(oObjDef) != null) {
        return g_datePeriodCache.get(oObjDef);
    }
    
    var aLanes = oModel.GetLanes(Constants.LANE_VERTICAL);
    var oOccLane = null;
    for (var i=0; i<aLanes.length; i++) {
        if (aLanes[i].IsObjOccOfLane(oObjOcc)) {
            oOccLane = aLanes[i];
            break;
        }
    }
    if (oOccLane == null) {
        return;
    }
    
    var iStartdateAttributeTypeNum = oOccLane.Attribute(Constants.AT_POSITION_ATTR, g_nLoc).MeasureValue();   //Date attribute       - 369 - Attribut AT_DATE_START ("Start date")
    var iDurationAttributeTypeNum     = oOccLane.Attribute(Constants.AT_SIZE_ATTR, g_nLoc).MeasureValue();       //Duration attribute   - 718 - Attribut AT_MAX_TL_TIME ("Maximum total time")
    
    var oStartDate = oObjDef.Attribute(iStartdateAttributeTypeNum, g_nLoc).MeasureValue();   // Date object
    var oEndDate = oObjDef.getEndDate(iStartdateAttributeTypeNum, iDurationAttributeTypeNum, g_nLoc);   // Calculated Date object
    
    var datePeriodContainer = new DatePeriodContainer();
    datePeriodContainer.startDate = oStartDate;
    datePeriodContainer.endDate = oEndDate;
    
    //update cache
    g_datePeriodCache.put(oObjDef, datePeriodContainer);
    
    return datePeriodContainer;
}

function calculateEndDate(startDate, duration, durationUnit) {
    
    var calendar = java.util.Calendar.getInstance();
    calendar.setTime(startDate);
    
    var millis = duration;
    switch (durationUnit) {  
        case Constants.AVT_YEARS:
            millis = millis * 12; 
        case Constants.AVT_MON:
            millis = millis * 365;
        case Constants.AVT_DAYS:
            millis = millis * 24; 
        case Constants.AVT_HRS:
            millis = millis * 60;
        case Constants.AVT_MINS:
            millis = millis * 60;
        case Constants.AVT_SECS:
            millis = millis * 1000;
    }
    var millisSum = calendar.getTimeInMillis() + millis;
    calendar.setTimeInMillis(millisSum);
    return calendar.getTime();
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

    for(var i=0; i<Math.min(tokenList.size(), 4); i++) {
        timeInMillis += java.lang.Long.parseLong(tokenList.get(i)) * mults[i];
    }

    return timeInMillis * 1000;    
}