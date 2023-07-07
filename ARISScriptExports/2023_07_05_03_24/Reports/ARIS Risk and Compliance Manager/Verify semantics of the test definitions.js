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
   
var gRelevantModelsSet = new java.util.HashSet(); // needed for uniqueness check of testdefintions

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
    var aTestdefinitions = convertHashSetToJSArray(g_testdefinitionSet);
    
    
    //3. checks
    for (var i=0; i<aTestdefinitions.length; i++){
        
        var szInnerMsg = "";     
        //Test management checks
        if (g_bTestManagementExport) {
            szInnerMsg = addSingleValidationOutput( validateTestdefinitionUniquenessOverBCDs(aTestdefinitions[i]), szInnerMsg );
            szInnerMsg = addSingleValidationOutput( validateTestdefinitionMandatoryAttributes(aTestdefinitions[i]), szInnerMsg );
            szInnerMsg = addSingleValidationOutput( validateTestdefinitionStartEndDate(aTestdefinitions[i]), szInnerMsg );
            szInnerMsg = addSingleValidationOutput( validateTestdefinitionFrequency(aTestdefinitions[i]), szInnerMsg );
            //deactivated for GW, CABI - 27.3.12
			//szInnerMsg = addSingleValidationOutput( validateTestdefinitionType(aTestdefinitions[i]), szInnerMsg );
            szInnerMsg = addSingleValidationOutput( validateConnectedOrgaUnitsCount(aTestdefinitions[i]), szInnerMsg );
            szInnerMsg = addSingleValidationOutput( validateConnectedTestmanagerCount(aTestdefinitions[i]), szInnerMsg );
            szInnerMsg = addSingleValidationOutput( validateConnectedTesterCount(aTestdefinitions[i]), szInnerMsg );
            szInnerMsg = addSingleValidationOutput( validateConnectedTestReviewerCount(aTestdefinitions[i]), szInnerMsg );
            szInnerMsg = addSingleValidationOutput( validateConnectedUsers(aTestdefinitions[i]), szInnerMsg );
            szInnerMsg = addSingleValidationOutput( validateConnectedControlsCount(aTestdefinitions[i]), szInnerMsg );
            //additional check
            if (TESTER_GROUPS_ASSIGNMENT_TESTER_HIERARCHY_PERFORM == true) {
                szInnerMsg = addSingleValidationOutput( validateConnectedTesterGroupHierarchyAssignment(aTestdefinitions[i]), szInnerMsg );
            }
        }
         //Risk management checks
        if (g_bRiskManagementExport) {
            //No test definition specific checks for risk management so far
        }

        //in case of error add the test definition object info
        if (!szInnerMsg.equals("")){
            
            var szTestDefInfo = new java.lang.String(getString("TEXT_2")).replaceFirst("%0", aTestdefinitions[i].Name(g_nLoc));
            var sCompleteSingleTestDefOutput = addObjectValidationInfo(szTestDefInfo, szInnerMsg, aTestdefinitions[i], SPC1)
            szOutput = addCompleteObjectValidationOutput(szOutput, sCompleteSingleTestDefOutput); 
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
//-------------------------- Test management validates --------------------------------
function validateTestdefinitionMandatoryAttributes(pTestdefinitions){
    var szMsg = "";
    var aMandatories = getMandatoryFieldsForTestdefinition();
    return validateMandatoryObjectAttributes(pTestdefinitions, aMandatories, SPC1, getString("TEXT_3")); 
}

function validateTestdefinitionStartEndDate(pTestdefinition) {
	return validateStartEndDate(pTestdefinition, Constants.AT_AAM_TESTDEF_START_DATE, Constants.AT_AAM_TESTDEF_END_DATE, SPC1);
}

function validateTestdefinitionFrequency(pTestdefinition) {
    var szMsg = "";
    //if flag "event driven generation allowed" is set to "true" then the frequency must be "event driven"...
    if ( (isboolattributetrue(pTestdefinition, Constants.AT_EVENT_DRIVEN_TESTS_ALLOWED, g_nLoc) 
         && pTestdefinition.Attribute(Constants.AT_AAM_TEST_FREQUENCY, g_nLoc).MeasureUnitTypeNum() != Constants.AVT_AAM_TEST_FREQUENCY_EVENT_DRIVEN) 
         ) { 
        szMsg = SPC1 + getString("TEXT_11");
    }
    //... and vice versa!
    if ( (isboolattributefalse(pTestdefinition, Constants.AT_EVENT_DRIVEN_TESTS_ALLOWED, g_nLoc) 
         && pTestdefinition.Attribute(Constants.AT_AAM_TEST_FREQUENCY, g_nLoc).MeasureUnitTypeNum() == Constants.AVT_AAM_TEST_FREQUENCY_EVENT_DRIVEN)
         ) { 
        szMsg = SPC1 + getString("TEXT_12");
    }
    return szMsg;
}


function validateTestdefinitionType(pTestdefinition) {
    var szMsg = "";
    //if flag "event driven generation allowed" is set to "true" then the test type must be "Test of effectiveness"
    if ( isboolattributetrue(pTestdefinition, Constants.AT_EVENT_DRIVEN_TESTS_ALLOWED, g_nLoc)
         && ( isboolattributetrue(pTestdefinition, Constants.AT_AAM_TEST_TYPE_DESIGN, g_nLoc)
              || isboolattributefalse(pTestdefinition, Constants.AT_AAM_TEST_TYPE_EFFECTIVENESS, g_nLoc) )
              || !pTestdefinition.Attribute(Constants.AT_AAM_TEST_TYPE_EFFECTIVENESS, g_nLoc).IsMaintained() ) { 
        szMsg = SPC1 + getString("TEXT_13");
    }
    return szMsg;
}


function validateTestdefinitionUniquenessOverBCDs(pTestdefinition){

    //determine all BCDs of the testdefinition
    var aTestdefinitionObjOccs = pTestdefinition.OccList();
    for (var i=0; i<aTestdefinitionObjOccs.length; i++) {
        if ( aTestdefinitionObjOccs[i].Model().TypeNum() == Constants.MT_BUSY_CONTR_DGM ) {
            gRelevantModelsSet.add( aTestdefinitionObjOccs[i].Model() );
        }
    }
    
    var szMsg = validateObjectUniquenessOverModels( pTestdefinition, getString("TEXT_4"), convertHashSetToJSArray( gRelevantModelsSet ), SPC1 );
    return szMsg;
}


function validateConnectedOrgaUnitsCount(pTestdefinition){
    
   var aConnectedOrgaUnits = g_testdefinition2orgunitHashMap.get(pTestdefinition);
   return getConnectionCountValidationOutput(OCCURRENCE_MIN_TESTDEFINITION_TO_ORGAUNITS, OCCURRENCE_MAX_TESTDEFINITION_TO_ORGAUNITS, pTestdefinition, aConnectedOrgaUnits, getString("TEXT_5"), SPC1);     
    
}


function validateConnectedTestmanagerCount(pTestdefinition){
   
   var aConnectedGroups = g_testdefinition2testManagerHashMap.get(pTestdefinition);
   return getConnectionCountValidationOutput(OCCURRENCE_MIN_TESTDEFINITION_TO_TESTMANAGER, OCCURRENCE_MAX_TESTDEFINITION_TO_TESTMANAGER, pTestdefinition, aConnectedGroups, getString("TEXT_15"), SPC1);
}


function validateConnectedTesterCount(pTestdefinition){
   
   var aConnectedGroups = g_testdefinition2testerHashMap.get(pTestdefinition);
   return getConnectionCountValidationOutput(OCCURRENCE_MIN_TESTDEFINITION_TO_TESTER, OCCURRENCE_MAX_TESTDEFINITION_TO_TESTER, pTestdefinition, aConnectedGroups, getString("TEXT_6"), SPC1);
}


function validateConnectedTestReviewerCount(pTestdefinition){
   
   var aConnectedGroups = g_testdefinition2testReviewerHashMap.get(pTestdefinition);
   return getConnectionCountValidationOutput(OCCURRENCE_MIN_TESTDEFINITION_TO_TESTREVIEWER, OCCURRENCE_MAX_TESTDEFINITION_TO_TESTREVIEWER, pTestdefinition, aConnectedGroups, getString("TEXT_7"), SPC1);
}


/*---------------------------------------------------------------------------------------
    Checks the users which are connected with testgroup.
    It's not allowed to have the same user as tester and test reviewer
  ---------------------------------------------------------------------------------------*/
function validateConnectedUsers(pTestdefinition){
   
    var szMsg = "";
    //Connected tester and test reviewer groups
    var aConnectedTesterGroups = g_testdefinition2testerHashMap.get(pTestdefinition);
    var aConnectedTestReviewerGroups = g_testdefinition2testReviewerHashMap.get(pTestdefinition);
   
    var aAssignedGroupsSet = new java.util.HashSet();
    aAssignedGroupsSet.addAll( convertJSArrayToHashSet(aConnectedTesterGroups) );
    aAssignedGroupsSet.addAll( convertJSArrayToHashSet(aConnectedTestReviewerGroups) );

    aUsernames = haveGroupsSameUsers( convertHashSetToJSArray(aAssignedGroupsSet) );
    if (aUsernames.length > 0) {
        szMsg = SPC1 + getString("TEXT_10") + "\r\n";
        for (var i=0; i<aUsernames.length; i++) {
            if (i < aUsernames.length - 1) {szMsg += "\r\n";}
            szMsg += SPC1 + SPC1 + aUsernames[i];
        }
    } 
    return szMsg;
}

/*---------------------------------------------------------------------------------------
    Checks if both groups have same users assigned.
    HD: 12.2009  additional check for the tester group name
  ---------------------------------------------------------------------------------------*/   
function haveGroupsSameUsers(pConnectedGroups){
    
    var aUsernames = new Array();
    if (pConnectedGroups == null || pConnectedGroups.length < 2) {return aUsernames;}
    
    // relations to testers
    var cxnObjListGroup1 = pConnectedGroups[0].CxnListFilter(Constants.EDGES_IN, Constants.CT_EXEC_5);
    var cxnObjListGroup2 = pConnectedGroups[1].CxnListFilter(Constants.EDGES_IN, Constants.CT_EXEC_5);
    // check if there are the same objects related    
    for (var i=0; i<cxnObjListGroup1.length; i++){
        for(var j=0; j<cxnObjListGroup2.length; j++) {
           if ( cxnObjListGroup1[i].SourceObjDef().IsEqual(cxnObjListGroup2[j].SourceObjDef()) ){
                aUsernames.push( cxnObjListGroup2[j].SourceObjDef().Name(g_nLoc) );
           }
        }
    }
    return aUsernames;
}

/*---------------------------------------------------------------------------------------
    Checks if both group have different assigned roles - one is tester, one is test reviewer.
    Prerequisite is that each group is assigned to exactly one role.
  ---------------------------------------------------------------------------------------*/   
function haveGroupsTesterAndTestReviewerRoles(pConnectedGroups){

    //true if one group is tester and the other is reviewer
    return ( (isPersonTypeTester(pConnectedGroups[0]) && isPersonTypeTestReviewer(pConnectedGroups[1]))
             || (isPersonTypeTester(pConnectedGroups[1]) && isPersonTypeTestReviewer(pConnectedGroups[0])) );       
}


/*---------------------------------------------------------------------------------------
    Checks if at least one tester group is assigned.
    For all such tester groups it checks if they are assigned to one tester hierarchy 
    element only.
  ---------------------------------------------------------------------------------------*/
function validateConnectedTesterGroupHierarchyAssignment(pTestdefinition){
    
    var szMsg = "";
    var aConnectedGroups = g_testdefinition2testerHashMap.get(pTestdefinition);
    var aConnectedTesterGroups = new Array();
    
    //check if at least one assigned tester group is part of the tester hierarchy
    for (var i=0; i<aConnectedGroups.length; i++) {
        if (isPersonTypeInTestHierarchy(aConnectedGroups[i])) {
            aConnectedTesterGroups.push(aConnectedGroups[i]);
        }
    }
    if (aConnectedTesterGroups.length == 0) {
        szMsg += SPC1 + getString("TEXT_9");
        return szMsg;
    }
    
    //check if each assigned tester group is assigned to one tester hierarchy element only
    for (var i=0; i<aConnectedTesterGroups.length; i++) {
        var aAssignedTesterHierarchies = g_testerGroup2TesterHierarchyHashMap.get(aConnectedTesterGroups[i]);
        if (aAssignedTesterHierarchies.length > 1) {
            szMsg += SPC1 + new java.lang.String(getString("TEXT_14")).replaceFirst("%0", aConnectedTesterGroups[i].Name(gLanguage));
            for (var j=0; j<aAssignedTesterHierarchies.length; j++){
                szMsg += "\r\n" + SPC2 + aAssignedTesterHierarchies[j].Name(gLanguage);
            }
        }
    }
    
    return szMsg;
}

 
function validateConnectedControlsCount(pTestdefinition){
   
    //-> get all linked export relevant controls
    var aLinkedExportRelevantControls = g_testdefinition2controlsHashMap.get(pTestdefinition);
    //check cardinality
    return getConnectionCountValidationOutput(OCCURRENCE_MIN_TESTDEFINITION_TO_CONTROLS, OCCURRENCE_MAX_TESTDEFINITION_TO_CONTROLS, pTestdefinition, aLinkedExportRelevantControls, getString("TEXT_8"), SPC1); 
}