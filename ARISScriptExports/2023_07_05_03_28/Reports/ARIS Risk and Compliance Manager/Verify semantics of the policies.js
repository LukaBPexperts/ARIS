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
    if (!g_bPolicyManagementExport) {
        determineModules();    
    }
    //if no mode selected then nothing to do
    if (!g_bPolicyManagementExport) {return;}
    
    
    //2. selection
    getPolicyManagementSelectionByContext();
    var aPolicyDefinitions = convertHashSetToJSArray(g_policyDefinitionSet);
    
    
    //3. checks
    var aPolicyDefinitionMandatories = getMandatoryFieldsForPolicyDefinition();
	var aPolicyReviewTaskMandatories = getMandatoryFieldsForPolicyReviewTask();
    for (var i=0; i<aPolicyDefinitions.length; i++) {
        
        var szInnerMsg = "";
        
		//policy definition attributes
        szInnerMsg = addSingleValidationOutput( validateMandatoryObjectAttributes(aPolicyDefinitions[i], aPolicyDefinitionMandatories, SPC1), szInnerMsg );
		
		szInnerMsg = addSingleValidationOutput( validatePolicyOwnerApprovalStartEndDate(aPolicyDefinitions[i], SPC1), szInnerMsg );
		szInnerMsg = addSingleValidationOutput( validatePolicyApproverApprovalStartEndDate(aPolicyDefinitions[i], SPC1), szInnerMsg );
		szInnerMsg = addSingleValidationOutput( validatePolicyPublishingStartEndDate(aPolicyDefinitions[i], SPC1), szInnerMsg );
		szInnerMsg = addSingleValidationOutput( validatePolicyReviewTaskStartEndDate(aPolicyDefinitions[i], SPC1), szInnerMsg );
		
        szInnerMsg = addSingleValidationOutput( validateConnectedPolicyOwnerCount(aPolicyDefinitions[i], SPC1), szInnerMsg );
        szInnerMsg = addSingleValidationOutput( validateConnectedPolicyAuditorCount(aPolicyDefinitions[i], SPC1), szInnerMsg );
        szInnerMsg = addSingleValidationOutput( validateConnectedPolicyApproverCount(aPolicyDefinitions[i], SPC1), szInnerMsg );
        
        szInnerMsg = addSingleValidationOutput( validatePolicyTypeConfirmationRequiredConditions(aPolicyDefinitions[i], SPC1), szInnerMsg );
        
        szInnerMsg = addSingleValidationOutput( validatePublishingStartdateBeforeOwnerApprovalStartdate(aPolicyDefinitions[i], SPC1), szInnerMsg );
        szInnerMsg = addSingleValidationOutput( validateOwnerApprovalPeriodWithinApproverApprovalPeriod(aPolicyDefinitions[i], SPC1), szInnerMsg );
        szInnerMsg = addSingleValidationOutput( validateDatePeriod(aPolicyDefinitions[i], Constants.AT_START_DATE_APPROVAL_PERIOD_OWNER, Constants.AT_END_DATE_APPROVAL_PERIOD_OWNER, SPC1), szInnerMsg );
        szInnerMsg = addSingleValidationOutput( validateDatePeriod(aPolicyDefinitions[i], Constants.AT_START_DATE_APPROVAL_PERIOD_APPROVER, Constants.AT_END_DATE_APPROVAL_PERIOD_APPROVER, SPC1), szInnerMsg );
        szInnerMsg = addSingleValidationOutput( validateDatePeriod(aPolicyDefinitions[i], Constants.AT_START_DATE_PUBLISHING_PERIOD, Constants.AT_END_DATE_PUBLISHING_PERIOD, SPC1), szInnerMsg );
        
		//policy review task attriutes
		szInnerMsg = addSingleValidationOutput( validateMandatoryObjectAttributes(aPolicyDefinitions[i], aPolicyReviewTaskMandatories, SPC1), szInnerMsg );
		
		
        //---- add header in case of errors
        if (!szInnerMsg.equals("")) {
            var szPolicyDefinitionInfo = new java.lang.String(getString("TEXT_2")).replaceFirst("%0", aPolicyDefinitions[i].Name(g_nLoc));
            var sCompleteSinglePolicyDefinitionOutput = addObjectValidationInfo(szPolicyDefinitionInfo, szInnerMsg, aPolicyDefinitions[i], SPC1);
            szOutput = addCompleteObjectValidationOutput(szOutput, sCompleteSinglePolicyDefinitionOutput);
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

function validatePolicyOwnerApprovalStartEndDate(p_policyDefinition, p_sSPC) {
	return validateStartEndDate(p_policyDefinition, Constants.AT_START_DATE_APPROVAL_PERIOD_OWNER, Constants.AT_END_DATE_APPROVAL_PERIOD_OWNER, p_sSPC);
}

function validatePolicyApproverApprovalStartEndDate(p_policyDefinition, p_sSPC) {
	return validateStartEndDate(p_policyDefinition, Constants.AT_START_DATE_APPROVAL_PERIOD_APPROVER, Constants.AT_END_DATE_APPROVAL_PERIOD_APPROVER, p_sSPC);
}

function validatePolicyPublishingStartEndDate(p_policyDefinition, p_sSPC) {
	return validateStartEndDate(p_policyDefinition, Constants.AT_START_DATE_PUBLISHING_PERIOD, Constants.AT_END_DATE_PUBLISHING_PERIOD, p_sSPC);
}

function validatePolicyReviewTaskStartEndDate(p_policyDefinition, p_sSPC) {
	return validateStartEndDate(p_policyDefinition, Constants.AT_START_DATE_OF_POLICY_REVIEWS, Constants.AT_END_DATE_OF_POLICY_REVIEWS, p_sSPC);
}

function validateConnectedPolicyOwnerCount(p_policyDefinition, p_sSPC) {
   
   var aConnectedGroups = g_policyDefinition2policyOwnerHashMap.get(p_policyDefinition);
   return getConnectionCountValidationOutput(PM_OCCURRENCE_MIN_POLICY_DEFINITION_TO_POLICYOWNER, PM_OCCURRENCE_MAX_POLICY_DEFINITION_TO_POLICYOWNER, p_policyDefinition, aConnectedGroups, getString("TEXT_3"), SPC1);
}

function validateConnectedPolicyAuditorCount(p_policyDefinition, p_sSPC) {
   
   var aConnectedGroups = g_policyDefinition2policyAuditorHashMap.get(p_policyDefinition);
   return getConnectionCountValidationOutput(PM_OCCURRENCE_MIN_POLICY_DEFINITION_TO_POLICYAUDITOR, PM_OCCURRENCE_MAX_POLICY_DEFINITION_TO_POLICYAUDITOR, p_policyDefinition, aConnectedGroups, getString("TEXT_9"), SPC1);
}

function validateConnectedPolicyApproverCount(p_policyDefinition, p_sSPC) {
   
   var aConnectedGroups = g_policyDefinition2policyApproverHashMap.get(p_policyDefinition);
   return getConnectionCountValidationOutput(PM_OCCURRENCE_MIN_POLICY_DEFINITION_TO_POLICYAPPROVER, PM_OCCURRENCE_MAX_POLICY_DEFINITION_TO_POLICYAPPROVER, p_policyDefinition, aConnectedGroups, getString("TEXT_4"), SPC1);
}

/*
 These conditions are checked hard-coded here: if policy type "confirmation required"
 1. At least one Policy Addressee group must be assigned
 2.	Confirmation text must be maintained
 3.	Duration must be maintained (Range 1..365)

*/
function validatePolicyTypeConfirmationRequiredConditions(p_policyDefinition, p_sSPC) {

    var szMsg = "";

    //nothing to check here if attribute type is not "Confirmation required"
    if ( p_policyDefinition.Attribute(Constants.AT_POLICY_TYPE, g_nLoc).MeasureUnitTypeNum() != Constants.AVT_CONFIRMATION_REQUIRED ) {
        return szMsg;
    }
   
    var bValid = true;
    var sSubConditionsMsg = "";
    var sSingleCheckResult = "";
    
    //1.
    var aConnectedGroups = g_policyDefinition2policyAddresseeHashMap.get(p_policyDefinition);
    sSingleCheckResult = getConnectionCountValidationOutput(PM_OCCURRENCE_MIN_POLICY_DEFINITION_TO_POLICYADDRESSEE, PM_OCCURRENCE_MAX_POLICY_DEFINITION_TO_POLICYADDRESSEE, p_policyDefinition, aConnectedGroups, getString("TEXT_5"), p_sSPC);
    if (szMsg.length > 0) {szMsg += "\r\n";}
    szMsg += sSingleCheckResult;
       
    //2.
    aMandatories = new Array( new Array("AT_CONFIRMATION_DURATION[Range 1...]") );
    sSingleCheckResult = validateMandatoryObjectAttributes(p_policyDefinition, aMandatories, p_sSPC + SPC1, null);
    szMsg = addSingleValidationOutput(sSingleCheckResult, szMsg);
    
    
    //construct error message if condition(s) are violated
    if (szMsg.length > 0) {
       var sConfirmationRequiredItemName = ArisData.ActiveFilter().AttrValueType(Constants.AT_POLICY_TYPE, Constants.AVT_CONFIRMATION_REQUIRED);
       szMsg = p_sSPC + new java.lang.String(getString("TEXT_10")).replaceFirst("%0", sConfirmationRequiredItemName)
               + "\r\n" + "\r\n" + p_sSPC + szMsg;
    }
    
    return szMsg;        
}


function validatePublishingStartdateBeforeOwnerApprovalStartdate(p_policyDefinition, p_sSPC) {
    
    var szMsg = "";
    var ownerApprovalStartdateAttribute = p_policyDefinition.Attribute(Constants.AT_START_DATE_APPROVAL_PERIOD_OWNER, g_nLoc);
    var publishingStartdateAttribute = p_policyDefinition.Attribute(Constants.AT_START_DATE_PUBLISHING_PERIOD, g_nLoc);
    if (!ownerApprovalStartdateAttribute.IsMaintained() || !publishingStartdateAttribute.IsMaintained()) {
        return szMsg;
    }
    
    var ownerApprovalStartdate = ownerApprovalStartdateAttribute.MeasureValue();
    var publishingStartdate = publishingStartdateAttribute.MeasureValue();
    if (publishingStartdate.compareTo(ownerApprovalStartdate) < 0) {
        szMsg = p_sSPC + getString("TEXT_6");
    }

    return szMsg;
}

function validateOwnerApprovalPeriodWithinApproverApprovalPeriod(p_policyDefinition, p_sSPC) {
    
    var szMsg = "";
    var ownerApprovalStartdateAttribute = p_policyDefinition.Attribute(Constants.AT_START_DATE_APPROVAL_PERIOD_OWNER, g_nLoc);
    var ownerApprovalEnddateAttribute = p_policyDefinition.Attribute(Constants.AT_END_DATE_APPROVAL_PERIOD_OWNER, g_nLoc);
    var approverApprovalStartdateAttribute = p_policyDefinition.Attribute(Constants.AT_START_DATE_APPROVAL_PERIOD_APPROVER, g_nLoc);
    var approverApprovalEnddateAttribute = p_policyDefinition.Attribute(Constants.AT_END_DATE_APPROVAL_PERIOD_APPROVER, g_nLoc);
    if (!ownerApprovalStartdateAttribute.IsMaintained() || !ownerApprovalEnddateAttribute.IsMaintained()
        || !approverApprovalStartdateAttribute.IsMaintained() || !approverApprovalEnddateAttribute.IsMaintained()) {
        return szMsg;
    }
    
    var ownerApprovalStartdate = ownerApprovalStartdateAttribute.MeasureValue();
    var ownerApprovalEnddate = ownerApprovalEnddateAttribute.MeasureValue();
    var approverApprovalStartdate = approverApprovalStartdateAttribute.MeasureValue();
    var approverApprovalEnddate = approverApprovalEnddateAttribute.MeasureValue();
    if (approverApprovalStartdate.compareTo(ownerApprovalStartdate) < 0
        || approverApprovalEnddate.compareTo(ownerApprovalEnddate) > 0) {
        szMsg = p_sSPC + getString("TEXT_7");
    }
    
    return szMsg;
}

function validateDatePeriod(p_oObjDef, p_iStartDateTypeNum, p_iEndDateTypeNum, p_sSPC) {
    
    var szMsg = "";
    var oStartdateAttribute = p_oObjDef.Attribute(p_iStartDateTypeNum, g_nLoc);
    var oEnddateAttribute = p_oObjDef.Attribute(p_iEndDateTypeNum, g_nLoc);
    if (!oStartdateAttribute.IsMaintained() || !oEnddateAttribute.IsMaintained()) {
        return szMsg;
    }
    
    var oStartdate = oStartdateAttribute.MeasureValue();
    var oEnddate = oEnddateAttribute.MeasureValue();
    if (oEnddate.compareTo(oStartdate) < 0) {
        var sDatePeriodMsg = new java.lang.String(getString("TEXT_8") )
                                          .replaceFirst("%0", oEnddateAttribute.Type())
                                          .replaceFirst("%1", oStartdateAttribute.Type());
        szMsg = p_sSPC + sDatePeriodMsg;
    }
    
    return szMsg;
}