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

//Stores all error messages for the hierarchy elements; reused for all hierarchy types - Format: ObjDef | JSArray <String>
var g_hierarchyElement2Messages = new java.util.HashMap();
var g_alreadyProcessedHierarchyGUIDs = new java.util.HashSet();

var HIERARCHYTYPE_ORGUNIT                       = "HIERARCHYTYPE_ORGUNIT";
var HIERARCHYTYPE_ORGUNIT_RM					= "HIERARCHYTYPE_ORGUNIT_RM"; //for alternate mandatory attribute check in risk management
var HIERARCHYTYPE_ORGUNIT_CONNECTION            = Constants.CT_IS_SUPERIOR_1;
var HIERARCHYTYPE_ORGUNIT_CONNECTION_DIRECTION  = Constants.EDGES_IN;

var HIERARCHYTYPE_REGULATION                      = "HIERARCHYTYPE_REGULATION";
var HIERARCHYTYPE_REGULATION_RM					  = "HIERARCHYTYPE_REGULATION_RM"; //for alternate mandatory attribute check in risk management
var HIERARCHYTYPE_REGULATION_CONNECTION           = Constants.CT_HAS_1;
var HIERARCHYTYPE_REGULATION_CONNECTION_DIRECTION = Constants.EDGES_IN;

var HIERARCHYTYPE_PROCESS                       = "HIERARCHYTYPE_PROCESS";
var HIERARCHYTYPE_PROCESS_CONNECTION            = Constants.CT_IS_PRCS_ORNT_SUPER;
var HIERARCHYTYPE_PROCESS_CONNECTION_DIRECTION  = Constants.EDGES_IN | Constants.EDGES_ASSIGN;

var HIERARCHYTYPE_TESTER                        = "HIERARCHYTYPE_TESTER";
var HIERARCHYTYPE_TESTER_CONNECTION             = Constants.CT_IS_SUPERIOR_1;
var HIERARCHYTYPE_TESTER_CONNECTION_DIRECTION   = Constants.EDGES_IN;

var HIERARCHYTYPE_APPLICATIONSYSTEMTYPE                         = "HIERARCHYTYPE_APPLICATIONSYSTEMTYPE";
var HIERARCHYTYPE_APPLICATIONSYSTEMTYPE_CONNECTION              = Constants.CT_CAN_SUBS_2;
var HIERARCHYTYPE_APPLICATIONSYSTEMTYPE_CONNECTION_DIRECTION    = Constants.EDGES_IN;

var HIERARCHYTYPE_RISKCATEGORY                      = "HIERARCHYTYPE_RISKCATEGORY";
var HIERARCHYTYPE_RISKCATEGORY_CONNECTION           = Constants.CT_BELONG_CAT;
var HIERARCHYTYPE_RISKCATEGORY_CONNECTION_DIRECTION = Constants.EDGES_IN;

//---------------------------------------------------------------
              
var HEADER_REGULATION     		= getString("TEXT_2");
var HEADER_PROCESS              = getString("TEXT_3");
var HEADER_ORGANIZATION         = getString("TEXT_4");
var HEADER_TESTER               = getString("TEXT_5");
var HEADER_APPSYS               = getString("TEXT_6");
var HEADER_RISKCAT              = getString("TEXT_7");

var HEADER_ERROR_PARENT         = getString("TEXT_8");
var HEADER_ERROR_MANDATORY      = getString("TEXT_9");
var HEADER_ERROR_MULTIPLEGROUPS = getString("TEXT_10");
var HEADER_ERROR_CYCLE          = getString("TEXT_13");

//---------------------------------------------------------------

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
    
    
    //3. checks
    //Checks for both modes
    var sCheckResult = "";
    if (g_bTestManagementExport || g_bRiskManagementExport) {
    	
        //if there are organisational units that were classified as both orgunit hierarchy and tester hierarchy element
        //then create an error message here
        sCheckResult = writeBothOrgUnitTesterHierarchyAssignmentErrorObject("");
        szOutput = concatMessageStrings(szOutput, sCheckResult);
    }
    //Test management checks
    if (g_bTestManagementExport) {
        
		sCheckResult = checkHierarchy(HIERARCHYTYPE_ORGUNIT, HIERARCHYTYPE_ORGUNIT_CONNECTION, HIERARCHYTYPE_ORGUNIT_CONNECTION_DIRECTION, SPC1);
		szOutput = concatMessageStrings(szOutput, sCheckResult);
        
        sCheckResult = checkHierarchy(HIERARCHYTYPE_REGULATION, HIERARCHYTYPE_REGULATION_CONNECTION, HIERARCHYTYPE_REGULATION_CONNECTION_DIRECTION, SPC1);
		szOutput = concatMessageStrings(szOutput, sCheckResult);
        
        sCheckResult = checkHierarchy(HIERARCHYTYPE_PROCESS, HIERARCHYTYPE_PROCESS_CONNECTION, HIERARCHYTYPE_PROCESS_CONNECTION_DIRECTION, SPC1);
		szOutput = concatMessageStrings(szOutput, sCheckResult);
        
        sCheckResult = checkHierarchy(HIERARCHYTYPE_TESTER, HIERARCHYTYPE_TESTER_CONNECTION, HIERARCHYTYPE_TESTER_CONNECTION_DIRECTION, SPC1);
        szOutput = concatMessageStrings(szOutput, sCheckResult);
    }
    //Risk management checks
    if (g_bRiskManagementExport) {

        //test management checks are more powerful so skip this step if also test management mode is selected
        if (!g_bTestManagementExport) {
			sCheckResult = checkHierarchy(HIERARCHYTYPE_REGULATION_RM, HIERARCHYTYPE_REGULATION_CONNECTION, HIERARCHYTYPE_REGULATION_CONNECTION_DIRECTION, SPC1);
            szOutput = concatMessageStrings(szOutput, sCheckResult);
        }
		//test management checks are more powerful so skip this step if also test management mode is selected
        if (!g_bTestManagementExport) {
			sCheckResult = checkHierarchy(HIERARCHYTYPE_ORGUNIT_RM, HIERARCHYTYPE_ORGUNIT_CONNECTION, HIERARCHYTYPE_ORGUNIT_CONNECTION_DIRECTION, SPC1);
            szOutput = concatMessageStrings(szOutput, sCheckResult);
        }
        
		sCheckResult = checkHierarchy(HIERARCHYTYPE_APPLICATIONSYSTEMTYPE, HIERARCHYTYPE_APPLICATIONSYSTEMTYPE_CONNECTION, HIERARCHYTYPE_APPLICATIONSYSTEMTYPE_CONNECTION_DIRECTION, SPC1);
		szOutput = concatMessageStrings(szOutput, sCheckResult);
        
        sCheckResult = checkHierarchy(HIERARCHYTYPE_RISKCATEGORY, HIERARCHYTYPE_RISKCATEGORY_CONNECTION, HIERARCHYTYPE_RISKCATEGORY_CONNECTION_DIRECTION, SPC1);
        szOutput = concatMessageStrings(szOutput, sCheckResult);
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

function concatMessageStrings(pOverallString, pPartialString) {
    if (pPartialString == null || pPartialString.length == 0) {return pOverallString;}
    if (pOverallString == null) {return pPartialString;}
    if (pOverallString.length > 0) {pOverallString += "\r\n" + "\r\n" + "\r\n";}
    return pOverallString + pPartialString;
}

function checkHierarchy(pHierarchyType, pHierarchyConnection, pHierarchyConnectionDirection, pSPC) {

	//reset messages
	g_hierarchyElement2Messages = new java.util.HashMap();

	g_alreadyProcessedHierarchyGUIDs = new java.util.HashSet();
	
	sHeaderString = "";
	var jHierarchySet = null;
	if (pHierarchyType == HIERARCHYTYPE_REGULATION || pHierarchyType == HIERARCHYTYPE_REGULATION_RM) {
		jHierarchySet = g_regulationSet;
		sHeaderString = HEADER_REGULATION;
	}
	if (pHierarchyType == HIERARCHYTYPE_ORGUNIT || pHierarchyType == HIERARCHYTYPE_ORGUNIT_RM) {
		jHierarchySet = g_orgUnitSet;
		sHeaderString = HEADER_ORGANIZATION;
	}
	if (pHierarchyType == HIERARCHYTYPE_PROCESS) {
		jHierarchySet = g_processSet;
		sHeaderString = HEADER_PROCESS;
	}
	if (pHierarchyType == HIERARCHYTYPE_TESTER) {
		jHierarchySet = g_testerSet;
		sHeaderString = HEADER_TESTER;
	}
	if (pHierarchyType == HIERARCHYTYPE_APPLICATIONSYSTEMTYPE) {
		jHierarchySet = g_appSysSet;
		sHeaderString = HEADER_APPSYS;
	}
	if (pHierarchyType == HIERARCHYTYPE_RISKCATEGORY) {
		jHierarchySet = g_riskCategorySet;
		sHeaderString = HEADER_RISKCAT;
	}	
	
	//check all elements and generate error messages
	aHierarchyElements = convertHashSetToJSArray(jHierarchySet);
	for (var i=0; i<aHierarchyElements.length; i++) {
		checkHierarchyElement(aHierarchyElements[i], pHierarchyType, pHierarchyConnection, pHierarchyConnectionDirection, new java.util.ArrayList(), pSPC);
	}
	
	//combine all element messages
	var sCompleteHierarchyMessage = "";
	//element messages
	for (var j=0; j<aHierarchyElements.length; j++) {
		var aElementMessages = g_hierarchyElement2Messages.get(aHierarchyElements[j]);
        if (aElementMessages == null || aElementMessages.length == 0) {continue;}
        
        var sCompleteElementMessageString = "";
        for (var k=0; k<aElementMessages.length; k++) {
            if (sCompleteElementMessageString != "") {sCompleteElementMessageString += "\r\n" + "\r\n";}
            if (k == 0) {sCompleteElementMessageString += "\r\n";}
            sCompleteElementMessageString += aElementMessages[k]; 
        }
        
        if (sCompleteHierarchyMessage != "") {sCompleteHierarchyMessage += "\r\n" + "\r\n";}
        sCompleteHierarchyMessage += sCompleteElementMessageString;
	}

    if (!(sCompleteHierarchyMessage.length == 0)) {
        sCompleteHierarchyMessage = sHeaderString + sCompleteHierarchyMessage;
    }
	
	return sCompleteHierarchyMessage;
}

function checkHierarchyElement(pHierarchyElement, pHierarchyType, pHierarchyConnection, pHierarchyConnectionDirection, pCycleList, pSPC) {
    
    var aHierarchyMessages = new Array();
    var bCycleDetected = false;
    //if already contained in cycle then generate a cycle error message
    if (pCycleList.contains(pHierarchyElement)) {
        
        var sCycleMessage = pSPC + HEADER_ERROR_CYCLE;
        var aCycleListArray = convertHashSetToJSArray(pCycleList);
        //start at the cycle beginning which must be the current pHierarchyElement
        while (!aCycleListArray[0].equals(pHierarchyElement)) {aCycleListArray.shift();}
        
        for (var i=0; i<aCycleListArray.length; i++) {
            sCycleMessage   += "\r\n" + pSPC + aCycleListArray[i].Type() + " - " + aCycleListArray[i].Name(g_nLoc);
            sCycleMessage   += "\r\n" + pSPC + SPC1 + getString("TEXT_11");
            sCycleMessage   += "\r\n" + pSPC + SPC2 + getString("TEXT_12") + " " + aCycleListArray[i].Group().Path(g_nLoc);                
        }
        aHierarchyMessages.push(sCycleMessage);
        bCycleDetected = true;
    }
    
    //skip already processed elements except there is a cycle
    if ( !bCycleDetected && g_alreadyProcessedHierarchyGUIDs.contains(pHierarchyElement.GUID()) ) {return;}
	
    //validate mandatory attributes
    var sAttributeMsg = validateMandatoryAttributes(pHierarchyElement, pHierarchyType, pSPC);
    if (sAttributeMsg.length > 0) {
        aHierarchyMessages.push(sAttributeMsg);
    }
    
	//change review start and end date (not maintained for other hierarchy types, therefore not checked by validateStartEndDate())
	var sStartEndDateMessage = validateStartEndDate(pHierarchyElement, Constants.AT_REVIEW_START_DATE, Constants.AT_REVIEW_END_DATE, pSPC);
	if (sStartEndDateMessage.length > 0) {
		aHierarchyMessages.push(sStartEndDateMessage);
	}
    
    //validate connected groups
	var sGroupMessage = checkConnectedGroups(pHierarchyElement, pHierarchyType, pSPC);
    if (sGroupMessage.length > 0) {
        aHierarchyMessages.push(sGroupMessage);
    }
	
    //check direct parents
	var aParents = getConnectedObjects(pHierarchyElement, pHierarchyConnectionDirection, pHierarchyConnection, pHierarchyElement.TypeNum())
	if (aParents.length > 1) {
        var sParentsMessage = pSPC + HEADER_ERROR_PARENT;
        for (var j=0; j<aParents.length; j++) {
            if (j > 0) {sParentsMessage += "\r\n";}
            sParentsMessage   += pSPC + SPC1 + aParents[j].Type() + " - " + aParents[j].Name(g_nLoc);
            sParentsMessage   += "\r\n" + pSPC + SPC2 + getString("TEXT_11");
            sParentsMessage   += "\r\n" + pSPC + SPC3 + getString("TEXT_12") + " " + aParents[j].Group().Path(g_nLoc);
        }
        aHierarchyMessages.push(sParentsMessage);
    }
    
    //add the message which hierarchy element is meant
    if (aHierarchyMessages.length > 0) {
               
        var sElementTitleMessage    = pSPC + pHierarchyElement.Type() + " - " + pHierarchyElement.Name(g_nLoc);
        sElementTitleMessage        += "\r\n" + pSPC + "----------------------------------------";
        sElementTitleMessage        += "\r\n" + pSPC + getString("TEXT_11");
        sElementTitleMessage        += "\r\n" + pSPC + SPC1 + getString("TEXT_12") + " " + pHierarchyElement.Group().Path(g_nLoc);

        aHierarchyMessages.unshift(sElementTitleMessage);
        g_hierarchyElement2Messages.put(pHierarchyElement, aHierarchyMessages);
    }
    
    //store error messages and cycle info before processing the parents
    pCycleList.add(pHierarchyElement);
    g_alreadyProcessedHierarchyGUIDs.add(pHierarchyElement.GUID());
    
    //no recursion in case of an detected cycle
    if (bCycleDetected) {return;}
    
    //recursion for all direct parents
    for (var k=0; k<aParents.length; k++) {
		var recursionCycleList = new java.util.ArrayList(pCycleList);
		checkHierarchyElement(aParents[k], pHierarchyType, pHierarchyConnection, pHierarchyConnectionDirection, recursionCycleList, pSPC);
	}
    
    
}

function checkConnectedGroups(pHierarchyElement, pHierarchyType, pSPC) {
    var sMessage = "";
    var sConnectionMessage = "";
    var aConnectedGroups;
    
    //always check hierarchy owner groups
    aConnectedGroups = g_hierarchy2HierarchyOwnerHashMap.get(pHierarchyElement);
	//for regulations where AT_REVIEW_RELEVANT is set to "true" there must be assigned at least one hierarchy owner group
	if ( pHierarchyType == HIERARCHYTYPE_REGULATION && isboolattributetrue(pHierarchyElement, Constants.AT_REVIEW_RELEVANT, g_nLoc) ) {
		sConnectionMessage = getConnectionCountValidationOutput(1, OCCURRENCE_MAX_HIERARCHYOWNERGROUPS_TO_HIERARCHY, pHierarchyElement, aConnectedGroups, getString("TEXT_18"), pSPC);
		sMessage = addSingleValidationOutput( sConnectionMessage, sMessage );
	}
	//in all other cases use the defined min value from semantic properties
	else {
		sConnectionMessage = getConnectionCountValidationOutput(OCCURRENCE_MIN_HIERARCHYOWNERGROUPS_TO_HIERARCHY, OCCURRENCE_MAX_HIERARCHYOWNERGROUPS_TO_HIERARCHY, pHierarchyElement, aConnectedGroups, getString("TEXT_18"), pSPC);
		sMessage = addSingleValidationOutput( sConnectionMessage, sMessage );
	}
    
    //always check assigned test auditor groups
    aConnectedGroups = g_hierarchy2TestAuditorHashMap.get(pHierarchyElement);
    sConnectionMessage = getConnectionCountValidationOutput(OCCURRENCE_MIN_TESTAUDITORGROUPS_TO_HIERARCHY, OCCURRENCE_MAX_TESTAUDITORGROUPS_TO_HIERARCHY, pHierarchyElement, aConnectedGroups, getString("TEXT_15"), pSPC);
    sMessage = addSingleValidationOutput( sConnectionMessage, sMessage );

    //always check signoff owner groups except for risk category hierarchy and application systme type hierarchy
    if (pHierarchyType != HIERARCHYTYPE_RISKCATEGORY && pHierarchyType != HIERARCHYTYPE_APPLICATIONSYSTEMTYPE) {
        aConnectedGroups = g_hierarchy2SignoffOwnerHashMap.get(pHierarchyElement);
        sConnectionMessage = getConnectionCountValidationOutput(OCCURRENCE_MIN_SIGNOFFOWNERGROUPS_TO_HIERARCHY, OCCURRENCE_MAX_SIGNOFFOWNERGROUPS_TO_HIERARCHY, pHierarchyElement, aConnectedGroups, getString("TEXT_16"), pSPC);
        sMessage = addSingleValidationOutput( sConnectionMessage, sMessage );
    }
    //check assigned tester groups only for tester hierarchy
    if (pHierarchyType == HIERARCHYTYPE_TESTER) {
        aConnectedGroups = g_hierarchy2TesterGroupHashMap.get(pHierarchyElement);
        sConnectionMessage = getConnectionCountValidationOutput(OCCURRENCE_MIN_TESTERGROUPS_TO_HIERARCHY, OCCURRENCE_MAX_TESTERGROUPS_TO_HIERARCHY, pHierarchyElement, aConnectedGroups, getString("TEXT_17"), pSPC);
        sMessage = addSingleValidationOutput( sConnectionMessage, sMessage );
    }

	return sMessage;
}

function validateMandatoryAttributes(pObject, pHierarchyType, pSPC){
    var szMsg = "";
    var aMandatories = null;
     //Test management
    if (pHierarchyType.equals("HIERARCHYTYPE_REGULATION")){
        aMandatories = getMandatoryFieldsForHierarchyRegulation();   
    }
    if (pHierarchyType.equals("HIERARCHYTYPE_PROCESS")){
        aMandatories = getMandatoryFieldsForHierarchyProcess();
    }
    else if (pHierarchyType.equals("HIERARCHYTYPE_ORGUNIT")){
        aMandatories = getMandatoryFieldsForHierarchyOrganization(); 
    }
    else if (pHierarchyType.equals("HIERARCHYTYPE_TESTER")){
        aMandatories = getMandatoryFieldsForHierarchyTester();  
    }
    //Risk management
    else if (pHierarchyType.equals("HIERARCHYTYPE_REGULATION_RM")){
        aMandatories = getMandatoryRiskManagementFieldsForHierarchyRegulation();   
    }
    else if (pHierarchyType.equals("HIERARCHYTYPE_ORGUNIT_RM")){
        aMandatories = getMandatoryRiskManagementFieldsForHierarchyOrganization();   
    }
    else if (pHierarchyType.equals("HIERARCHYTYPE_APPLICATIONSYSTEMTYPE")){
        aMandatories = getMandatoryFieldsForHierarchyAppSys();  
    }
    else if (pHierarchyType.equals("HIERARCHYTYPE_RISKCATEGORY")){
        aMandatories = getMandatoryFieldsForHierarchyRiskCat();  
    }
    
    if (aMandatories != null) {       
        szMsg = validateMandatoryObjectAttributes(pObject, aMandatories, SPC1, HEADER_ERROR_MANDATORY);
    }
    return szMsg; 
}

//generate error message for this special case
function writeBothOrgUnitTesterHierarchyAssignmentErrorObject(pSPC) {

    szOut = "";
    var hierarchyIntersection = new java.util.HashSet(g_orgUnitSet);
    hierarchyIntersection.retainAll(g_testerSet);
    if (hierarchyIntersection.size() > 0) {
        var aErroneousHierarchyElements = convertHashSetToJSArray(hierarchyIntersection);
        szOut = getString("TEXT_14");
        for (var i=0; i<aErroneousHierarchyElements.length; i++) {
            if (i > 0) {szOut += "\r\n";}
            szOut += pSPC + SPC1 + "- " + aErroneousHierarchyElements[i].Name(g_nLoc);
        }
    }
    
    return szOut;
}