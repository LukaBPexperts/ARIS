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
    if (!g_bSurveyManagementExport) {
        determineModules();    
    }
    //if no mode selected then nothing to do
    if (!g_bSurveyManagementExport) {return;}
    
    
    //2. selection
    getSurveyManagementSelectionByContext();
    determineImplicitVisibilityDependencies();
    
    var aQuestionnaireTemplates = convertHashSetToJSArray(g_questionnaireTemplateSet);
    
    //3. checks
    for (var i=0; i<aQuestionnaireTemplates.length; i++) {
        
        var szInnerMsg = "";
        
        //---- questionnaire template itself
        var sQuestionnaireTemplateErrors = validateQuestionnaireTemplate(aQuestionnaireTemplates[i], SPC1);
        if (sQuestionnaireTemplateErrors != null && sQuestionnaireTemplateErrors.length > 0) {
            if (szInnerMsg.length > 0) {szInnerMsg += "\r\n" + "\r\n";}
            szInnerMsg += sQuestionnaireTemplateErrors;
        }
        
        //---- all sections and subsections
        var oFoundSectionsSet = new Packages.java.util.HashSet();
        var oFoundQuestionsSet = new Packages.java.util.HashSet();
        
        //determine all sections and questions beloning to the template directly or indirectly
        //search for section cycles
        var sOverallSectionStructureErrors = validateSectionStructure(aQuestionnaireTemplates[i], oFoundSectionsSet, oFoundQuestionsSet, SPC1);
        
        //now the individual section checks
        for (var it = oFoundSectionsSet.iterator(); it.hasNext();) {
            var oSection = it.next();
            
            var sIndiviualCombinedSectionErrors = "";
           
            //do a check on linked subordinated objects and mandatory fields for each single section
            var sSectionErrors = validateSectionSubordinatesAndMandatoryFields(oSection, SPC1);
            if (sSectionErrors != null && sSectionErrors.length > 0) {
                if (sIndiviualCombinedSectionErrors.length > 0) {sIndiviualCombinedSectionErrors += "\r\n" + "\r\n";}
                sIndiviualCombinedSectionErrors += sSectionErrors;
            }
                      
            //in case of errors at assigned subordinates or mandatory fields:
            //add also header with the section name
            if ( (sIndiviualCombinedSectionErrors != null && sIndiviualCombinedSectionErrors.length > 0) ) {
                var sSectionHeader = getString("TEXT_4") + " \"" + oSection.Name(g_nLoc) + "\":";
                sIndiviualCombinedSectionErrors = SPC1 + sSectionHeader + "\r\n"
                                                  + getUnderline(sSectionHeader, SPC1) + "\r\n"
                                                  + sIndiviualCombinedSectionErrors;
                
                if (szInnerMsg.length > 0) {szInnerMsg += "\r\n" + "\r\n"}
                szInnerMsg += sIndiviualCombinedSectionErrors;
            }
            
        }
        
        
        //for all sections: output the cycle errors only if there are no erroneous parent assignments
        //(because the cycle messages would be nonsense if the parents are incorrect)
        if ( (sOverallSectionStructureErrors != null && sOverallSectionStructureErrors.length > 0) ) {
            if (szInnerMsg.length > 0) {szInnerMsg += "\r\n" + "\r\n";}
            szInnerMsg += sOverallSectionStructureErrors;
        }
        
        
        //---- all questions   
        var sCombinedQuestionErrors = "";
        for (var it = oFoundQuestionsSet.iterator(); it.hasNext();) {
            var oQuestion = it.next();
            var sQuestionErrors = validateQuestion(oQuestion, SPC1);
            if (sQuestionErrors != null && sQuestionErrors.length > 0) {
                if (sCombinedQuestionErrors.length > 0) {sCombinedQuestionErrors += "\r\n" + "\r\n";}
                sCombinedQuestionErrors += sQuestionErrors;
            }
        }
        if (sCombinedQuestionErrors.length > 0) {
            if (szInnerMsg.length > 0) {szInnerMsg += "\r\n" + "\r\n";}
            szInnerMsg += sCombinedQuestionErrors;
        }
        
        //---- all option sets
        var oFoundOptionSetsSet = new Packages.java.util.HashSet();
        for (var it = oFoundQuestionsSet.iterator(); it.hasNext();) {
            var oQuestion = it.next();
            if ( g_question2optionSetsHashMap.get(oQuestion) != null ) {
                var aLinkedOptionSets = g_question2optionSetsHashMap.get(oQuestion);
                for (var j=0; j<aLinkedOptionSets.length; j++) {
                    oFoundOptionSetsSet.add( aLinkedOptionSets[j] );  
                }
            }  
        }
        var sCombinedOptionSetErrors = "";
        for (var it = oFoundOptionSetsSet.iterator(); it.hasNext();) {
            var oOptionSet = it.next();
            var sOptionSetErrors = validateOptionSet(oOptionSet, SPC1);
            if (sOptionSetErrors != null && sOptionSetErrors.length > 0) {
                if (sCombinedOptionSetErrors.length > 0) {sCombinedOptionSetErrors += "\r\n" + "\r\n";}
                sCombinedOptionSetErrors += sOptionSetErrors;
            }
        }
        if (sCombinedOptionSetErrors.length > 0) {
            if (szInnerMsg.length > 0) {szInnerMsg += "\r\n" + "\r\n";}
            szInnerMsg += sCombinedOptionSetErrors;
        }
        
        //---- all options
        var oFoundOptionsSet = new Packages.java.util.HashSet();
        for (var it = oFoundOptionSetsSet.iterator(); it.hasNext();) {
            var oOptionSet = it.next();
            if ( g_optionSet2optionsHashMap.get(oOptionSet) != null ) {
                var aLinkedOptions = g_optionSet2optionsHashMap.get(oOptionSet);
                for (var j=0; j<aLinkedOptions.length; j++) {
                    oFoundOptionsSet.add( aLinkedOptions[j] );  
                } 
            }  
        }
        for (var it = oFoundQuestionsSet.iterator(); it.hasNext();) {
            var oQuestion = it.next();
            if ( g_question2optionsHashMap.get(oQuestion) != null ) {
                var aLinkedOptions = g_question2optionsHashMap.get(oQuestion);
                for (var j=0; j<aLinkedOptions.length; j++) {
                    oFoundOptionsSet.add( aLinkedOptions[j] );  
                }
            }  
        }
        
        var sCombinedOptionErrors = "";
        for (var it = oFoundOptionsSet.iterator(); it.hasNext();) {
            var oOption = it.next();
            var sOptionErrors = validateOption(oOption, SPC1);
            if (sOptionErrors != null && sOptionErrors.length > 0) {
                if (sCombinedOptionErrors.length > 0) {sCombinedOptionErrors += "\r\n" + "\r\n";}
                sCombinedOptionErrors += sOptionErrors;
            }
        }
        if (sCombinedOptionErrors.length > 0) {
            if (szInnerMsg.length > 0) {szInnerMsg += "\r\n" + "\r\n";}
            szInnerMsg += sCombinedOptionErrors;
        }
        
		//---- all survey tasks
		var sCombinedSurveyTaskErrors = "";
		var aAssignedSurveyTasks = g_questionnaireTemplate2surveyTaskHashMap.get(aQuestionnaireTemplates[i]);
		for (var j=0; j<aAssignedSurveyTasks.length; j++) {
            var sSurveyTaskErrors = validateSurveyTask(aAssignedSurveyTasks[j], SPC1);
            if (sSurveyTaskErrors != null && sSurveyTaskErrors.length > 0) {
                if (sCombinedSurveyTaskErrors.length > 0) {sCombinedSurveyTaskErrors += "\r\n" + "\r\n";}
                sCombinedSurveyTaskErrors += sSurveyTaskErrors;
            }
        }
		if (sCombinedSurveyTaskErrors.length > 0) {
            if (szInnerMsg.length > 0) {szInnerMsg += "\r\n" + "\r\n";}
            szInnerMsg += sCombinedSurveyTaskErrors;
        }
		
        
        //---- add header in case of errors
        if (!szInnerMsg.equals("")){
            
            var szQuestionnaireTemplateInfo = new java.lang.String(getString("TEXT_2") ).replaceFirst("%0", aQuestionnaireTemplates[i].Name(g_nLoc));
            var sCompleteSingleQuestionnaireTemplateOutput = addObjectValidationInfo(szQuestionnaireTemplateInfo, szInnerMsg, aQuestionnaireTemplates[i], SPC1);
            szOutput = addCompleteObjectValidationOutput(szOutput, sCompleteSingleQuestionnaireTemplateOutput); 
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
function validateQuestionnaireTemplate(p_oQuestionnaireTemplate, p_sSPC) {
    
    var szMsg = "";
    
    //mandatory fields
    var sMandatoriesMsg = "";
    var aMandatories = getMandatoryFieldsForQuestionnaireTemplates();
    sMandatoriesMsg = validateMandatoryObjectAttributes(p_oQuestionnaireTemplate, aMandatories, p_sSPC);
        
    if (!sMandatoriesMsg.equals("")){
        szMsg = addSingleValidationOutput(sMandatoriesMsg, szMsg);
    }
    
    //linked sections
    var sLinkedSectionsMsg = "";
    var aConnectedSections = new Array();
    if (g_questionnaireTemplate2sectionsHashMap.get(p_oQuestionnaireTemplate) != null) {aConnectedSections = g_questionnaireTemplate2sectionsHashMap.get(p_oQuestionnaireTemplate);}
    sLinkedSectionsMsg = getConnectionCountValidationOutput(SM_OCCURRENCE_MIN_QUESTIONNAIRE_TEMPLATE_TO_SECTION, SM_OCCURRENCE_MAX_QUESTIONNAIRE_TEMPLATE_TO_SECTION, p_oQuestionnaireTemplate, aConnectedSections, getString("TEXT_5"), p_sSPC);      
    if (!sLinkedSectionsMsg.equals("")) {
        szMsg = addSingleValidationOutput(sLinkedSectionsMsg, szMsg);
    }
	
	//linked survey tasks
    var sLinkedSurveyTasksMsg = "";
    var aConnectedSurveyTasks = new Array();
    if (g_questionnaireTemplate2surveyTaskHashMap.get(p_oQuestionnaireTemplate) != null) {aConnectedSurveyTasks = g_questionnaireTemplate2surveyTaskHashMap.get(p_oQuestionnaireTemplate);}
    sLinkedSurveyTasksMsg = getConnectionCountValidationOutput(SM_OCCURRENCE_MIN_QUESTIONNAIRE_TEMPLATE_TO_SURVEYTASK, SM_OCCURRENCE_MAX_QUESTIONNAIRE_TEMPLATE_TO_SURVEYTASK, p_oQuestionnaireTemplate, aConnectedSurveyTasks, getString("TEXT_18"), p_sSPC);      
    if (!sLinkedSurveyTasksMsg.equals("")) {
        szMsg = addSingleValidationOutput(sLinkedSurveyTasksMsg, szMsg);
    } 

    if (szMsg.length > 0) {szMsg += "\r\n";}
    return szMsg;
}

//---------
//checks for cycles in the section structure
function validateSectionStructure(p_oQuestionnaireTemplate, p_oFoundSectionsSet, p_oFoundQuestions, p_sSPC) {

    //explicit sections and questions
    var aTopSections = g_questionnaireTemplate2sectionsHashMap.get(p_oQuestionnaireTemplate);
    
    if (aTopSections != null) {
        
        var oFoundSectionsSet = new Packages.java.util.HashSet();
        var oSectionStack = new Packages.java.util.Stack();
        var aSectionCycleErrors = new Array();
        for (var i=0; i<aTopSections.length; i++) {
            checkSectionCyclesRecursively(aTopSections[i], oFoundSectionsSet, p_oFoundQuestions, oSectionStack, aSectionCycleErrors);    
        }
        p_oFoundSectionsSet.addAll(oFoundSectionsSet);
    }
    
    //if cycle errors were found
    var sCycleErrorString = "";
    for (var i=0; i<aSectionCycleErrors.length; i++) {
        if (sCycleErrorString.length > 0) {sCycleErrorString += "\r\n";}
        sCycleErrorString += p_sSPC + aSectionCycleErrors[i];
    }
    if (sCycleErrorString.length > 0) {
        sCycleErrorString = p_sSPC + getString("TEXT_3") + "\r\n" + sCycleErrorString;
    }
    
    //implicit sections and questions
    var aImplicitSections = g_realImplicitSections.get(p_oQuestionnaireTemplate);
    for (var i=0; i<aImplicitSections.length; i++) {
        p_oFoundSectionsSet.add(aImplicitSections[i].oObjDefToPlace);
    }
    var aImplicitQuestions = g_realImplicitQuestions.get(p_oQuestionnaireTemplate);
    for (var j=0; j<aImplicitQuestions.length; j++) {
        p_oFoundQuestions.add(aImplicitQuestions[j].oObjDefToPlace);
    }
    
    return sCycleErrorString;    
}

//---------
function checkSectionCyclesRecursively(p_oSection, p_oFoundSectionsSet, p_oFoundQuestionSet, p_oStack, p_aSectionCycleErrors) {
    
    //if section is in the stack of already handled sections then we have a cycle
    if (p_oStack.contains(p_oSection)) {  
        var sCycleMessage = createCycleMessage(p_oSection, p_oStack, p_oFoundSectionsSet, p_oFoundQuestionSet);
        p_aSectionCycleErrors.push(SPC1 + sCycleMessage);
        return;
    }
    
    //Add it to the set of found sections...
    p_oFoundSectionsSet.add(p_oSection);
    
    //...put it on the stack...
    var oStack = new Packages.java.util.Stack();
    oStack.addAll(p_oStack);
    oStack.push(p_oSection);
    
    //...find all direct sub sections...
    var aSubSections = g_section2subsectionsHashMap.get(p_oSection);
    //...check all found sections recursively
    if (aSubSections != null) {
        for (var i=0; i<aSubSections.length; i++) {
            checkSectionCyclesRecursively(aSubSections[i], p_oFoundSectionsSet, p_oFoundQuestionSet, oStack, p_aSectionCycleErrors);    
        }
    }

    //...find all direct questions...   
    var aQuestions = g_section2questionsHashMap.get(p_oSection);
    //...check all found questions recursively
    for (var h=0; h<aQuestions.length; h++) { 
        checkSectionCycleQuestionsRecursively(aQuestions[h], p_oFoundSectionsSet, p_oFoundQuestionSet, oStack, p_aSectionCycleErrors);
    }
}

function checkSectionCycleQuestionsRecursively(p_oQuestion, p_oFoundSectionsSet, p_oFoundQuestionsSet, p_oStack, p_aSectionCycleErrors) {
    
    //if question is in the stack of already handled sections then we have a cycle
    if (p_oStack.contains(p_oQuestion)) {  
        var sCycleMessage = createCycleMessage(p_oQuestion, p_oStack, p_oFoundSectionsSet, p_oFoundQuestionsSet);
        p_aSectionCycleErrors.push(SPC1 + sCycleMessage);
        return;
    }
    
    //Add it to the set of found questions
    p_oFoundQuestionsSet.add(p_oQuestion);
}

function createCycleMessage(p_oCycleStartElement, p_oStack, p_oFoundSectionsSet, p_oFoundQuestions) {
    //make sure that the cycle text starts and ends with the same element
    p_oStack.push(p_oCycleStartElement);
    var sCycleMessage = "";
    var bCycleBeginningReached = false;
    for (var it = p_oStack.iterator(); it.hasNext(); ) {
        
        var oStackItem = it.next();
        bCycleBeginningReached = bCycleBeginningReached || oStackItem.equals(p_oCycleStartElement);
        if (!bCycleBeginningReached) {continue;}
        
        if (sCycleMessage.length > 0) {sCycleMessage += " - ";}
        var sObjType = "";
        if (p_oFoundSectionsSet.contains(oStackItem)) {sObjType = getString("TEXT_5");}
        if (p_oFoundQuestions.contains(oStackItem)) {sObjType = getString("TEXT_6");}
        sCycleMessage += " " + sObjType + " '" + oStackItem.Name(g_nLoc) + "'";
    } 
    return sCycleMessage;
}


//---------
function validateSectionSubordinatesAndMandatoryFields(p_oSection, p_sSPC) {
    
    var szMsg = "";
    
    //mandatory fields
    var sMandatoriesMsg = "";
    var aMandatories = getMandatoryFieldsForSections();
    sMandatoriesMsg = validateMandatoryObjectAttributes(p_oSection, aMandatories, p_sSPC);
        
    if (!sMandatoriesMsg.equals("")){
        szMsg = sMandatoriesMsg;
    }
    if (szMsg.length > 0) {szMsg += "\r\n";}
    
    //linked subordinated objects
    var sLinkedSectionsMsg = "";
    var aConnectedSections = new Array();
    if (g_section2subsectionsHashMap.get(p_oSection) != null) {aConnectedSections = g_section2subsectionsHashMap.get(p_oSection);}
    sLinkedSectionsMsg = getConnectionCountValidationOutput(SM_OCCURRENCE_MIN_SECTION_TO_SUBSECTION, SM_OCCURRENCE_MAX_SECTION_TO_SUBSECTION, p_oSection, aConnectedSections, getString("TEXT_5"), SPC1);      
    if (!sLinkedSectionsMsg.equals("")) {
        szMsg += sLinkedSectionsMsg;
    }
    if (szMsg.length > 0) {szMsg += "\r\n";}
    
    var sLinkedQuestionsMsg = "";
    var aConnectedQuestions = new Array();
    if (g_section2questionsHashMap.get(p_oSection) != null) {aConnectedQuestions = g_section2questionsHashMap.get(p_oSection);}
    sLinkedQuestionsMsg = getConnectionCountValidationOutput(SM_OCCURRENCE_MIN_SECTION_TO_QUESTION, SM_OCCURRENCE_MAX_SECTION_TO_QUESTION, p_oSection, aConnectedQuestions, getString("TEXT_7"), SPC1);      
    if (!sLinkedQuestionsMsg.equals("")) {
        szMsg += sLinkedQuestionsMsg;
    }
    if (szMsg.length > 0) {szMsg += "\r\n";}
    
    return szMsg;    
}

//---------
function validateQuestion(p_oQuestion, p_sSPC) {
    
    var szMsg = "";
    
    //---- mandatory fields
    var sMandatoriesMsg = "";
    var aMandatories = getMandatoryFieldsForQuestions();
    sMandatoriesMsg = validateMandatoryObjectAttributes(p_oQuestion, aMandatories, p_sSPC);
        
    if (!sMandatoriesMsg.equals("")){
        szMsg = sMandatoriesMsg;
    }
    
    //linked objects
    var sLinkedOptionSetsMsg = "";
    var aConnectedOptionSets = new Array();
    if (g_question2optionSetsHashMap.get(p_oQuestion) != null) {aConnectedOptionSets = g_question2optionSetsHashMap.get(p_oQuestion);}
    sLinkedOptionSetsMsg = getConnectionCountValidationOutput(SM_OCCURRENCE_MIN_QUESTION_TO_OPTIONSET, SM_OCCURRENCE_MAX_QUESTION_TO_OPTIONSET, p_oQuestion, aConnectedOptionSets, getString("TEXT_9"), SPC1);      
    if (!sLinkedOptionSetsMsg.equals("")) {
        if (szMsg != "") {szMsg += "\r\n" + "\r\n";}
        szMsg += sLinkedOptionSetsMsg;
    }

    var sLinkedOptionsMsg = "";
    var aConnectedOptions = new Array();
    if (g_question2optionsHashMap.get(p_oQuestion) != null) {aConnectedOptions = g_question2optionsHashMap.get(p_oQuestion);}
    sLinkedOptionsMsg = getConnectionCountValidationOutput(SM_OCCURRENCE_MIN_QUESTION_TO_OPTION, SM_OCCURRENCE_MAX_QUESTION_TO_OPTION, p_oQuestion, aConnectedOptions, getString("TEXT_11"), SPC1);      
    if (!sLinkedOptionsMsg.equals("")) {
        if (szMsg != "") {szMsg += "\r\n" + "\r\n";}
        szMsg += sLinkedOptionsMsg;
    }
    
    var oQuestionTypeAttribute = p_oQuestion.Attribute( Constants.AT_QUESTION_TYPE, g_nLoc);
    var iUnitTypeNum = oQuestionTypeAttribute.MeasureUnitTypeNum();
    var iSingleChoiceItemTypeNum =  Constants.AVT_SINGLE_CHOICE;
    var iMultipleChoiceItemTypeNum = Constants.AVT_MULTIPLE_CHOICE;
        
    //---- error if question is of type "single choice" or "multiple choice" but neither OptionSets nor Options are assigned 
    if ( (iUnitTypeNum == iSingleChoiceItemTypeNum || iUnitTypeNum == iMultipleChoiceItemTypeNum)
        && aConnectedOptionSets.length == 0 && aConnectedOptions.length == 0) {

        szMsg = addSingleValidationOutput(getString("TEXT_16"), szMsg, p_sSPC);
    }
    
    //---- error if both OptionSets and Options are assigned
    if (aConnectedOptionSets.length > 0 && aConnectedOptions.length > 0) {

        szMsg = addSingleValidationOutput(getString("TEXT_12"), szMsg, p_sSPC);
    }
    
    var iReviewerRatesAnswerTypeNum = Constants.AT_REVIEWER_RATES_ANSWER;
    
    if (oQuestionTypeAttribute.IsMaintained()) {
        
        if ( (aConnectedOptionSets.length > 0 || aConnectedOptions.length > 0) ) {
            
            //---- "reviewer rates answer" is not allowed to be true if the question is of type "single choice" or "multiple choice"
            if ( isSingleOrMultiOptionType(p_oQuestion)  
                && isboolattributetrue(p_oQuestion, iReviewerRatesAnswerTypeNum, g_nLoc) ) {

                szMsg = addSingleValidationOutput(getString("TEXT_13"), szMsg, p_sSPC);
            }
            
            //---- "reviewer rates answer" is not allowed to be false if the question type differs from "single choice" and "multiple choice"
            if ( !isSingleOrMultiOptionType(p_oQuestion)
                && !(isboolattributetrue(p_oQuestion, iReviewerRatesAnswerTypeNum, g_nLoc)) ) {
                
                szMsg = addSingleValidationOutput(getString("TEXT_13"), szMsg, p_sSPC);
            }
        } 
        else {
        
            //if there are neither option set nor options assigned then "reviewer rates answer" is not allowed to be "true"
            if (isboolattributetrue(p_oQuestion, iReviewerRatesAnswerTypeNum, g_nLoc)) {
                szMsg = addSingleValidationOutput(getString("TEXT_17"), szMsg, p_sSPC);
            }
        }
    }
    
    //in case of errors add header with the question text in AT_DESC; if not maintained then use AT_NAME as fallback
    if (szMsg.length > 0) {
        var sQuestionIdentifier = p_oQuestion.Attribute(Constants.AT_DESC, g_nLoc).getValue();
		if (!p_oQuestion.Attribute(Constants.AT_DESC, g_nLoc).IsMaintained()) {
			sQuestionIdentifier = p_oQuestion.Attribute(Constants.AT_NAME, g_nLoc).getValue();
		}
        var sQuestionHeader = getString("TEXT_6") + " \"" + sQuestionIdentifier + "\":";
        szMsg = p_sSPC + sQuestionHeader + "\r\n"
                + getUnderline(sQuestionHeader, p_sSPC) + "\r\n"
                + szMsg + "\r\n";
    }
    
    return szMsg;
}

function isSingleOrMultiOptionType(p_oQuestion) {
    var oQuestionTypeAttribute = p_oQuestion.Attribute( Constants.AT_QUESTION_TYPE, g_nLoc);
    var iUnitTypeNum = oQuestionTypeAttribute.MeasureUnitTypeNum();
    var iSingleChoiceItemTypeNum =  Constants.AVT_SINGLE_CHOICE;
    var iMultipleChoiceItemTypeNum = Constants.AVT_MULTIPLE_CHOICE;
    return iUnitTypeNum == iSingleChoiceItemTypeNum || iUnitTypeNum == iMultipleChoiceItemTypeNum;
}

//---------
function validateOptionSet(p_oOptionSet, p_sSPC) {
    
    var szMsg = "";
    
    //mandatory fields
    var sMandatoriesMsg = "";
    var aMandatories = getMandatoryFieldsForOptionSets();
    sMandatoriesMsg = validateMandatoryObjectAttributes(p_oOptionSet, aMandatories, p_sSPC);
        
    if (!sMandatoriesMsg.equals("")){
        szMsg = sMandatoriesMsg;
    }
    
    //linked objects
    var sLinkedOptionsMsg = "";
    var aConnectedOptions = new Array();
    if (g_optionSet2optionsHashMap.get(p_oOptionSet) != null) {aConnectedOptions = g_optionSet2optionsHashMap.get(p_oOptionSet);}
    sLinkedOptionsMsg = getConnectionCountValidationOutput(SM_OCCURRENCE_MIN_OPTIONSET_TO_OPTION, SM_OCCURRENCE_MAX_OPTIONSET_TO_OPTION, p_oOptionSet, aConnectedOptions, getString("TEXT_11"), SPC1);      
    if (!sLinkedOptionsMsg.equals("")) {
        if (szMsg != "") {szMsg += "\r\n" + "\r\n";}
        szMsg += sLinkedOptionsMsg;
    }
    
    //in case of errors add header with the option set name
    if (szMsg.length > 0) {
        var sOptionSetHeader = getString("TEXT_8") + " \"" + p_oOptionSet.Name(g_nLoc) + "\":";
        szMsg = p_sSPC + sOptionSetHeader + "\r\n"
                + getUnderline(sOptionSetHeader, p_sSPC) + "\r\n"
                + szMsg + "\r\n";
    }
    
    return szMsg;
}

//---------
function validateOption(p_oOption, p_sSPC) {
    
    var szMsg = "";
    
    //mandatory fields
    var sMandatoriesMsg = "";
    var aMandatories = getMandatoryFieldsForOptions();
    sMandatoriesMsg = validateMandatoryObjectAttributes(p_oOption, aMandatories, p_sSPC);
        
    if (!sMandatoriesMsg.equals("")){
        szMsg = sMandatoriesMsg;
    }
    
    //in case of errors add header with the option name
    if (szMsg.length > 0) {
        var sOptionHeader = getString("TEXT_10") + " \"" + p_oOption.Name(g_nLoc) + "\":";
        szMsg = p_sSPC + sOptionHeader + "\r\n"
                + getUnderline(sOptionHeader, p_sSPC) + "\r\n"
                + szMsg + "\r\n";
    }
    
    return szMsg;    
}

//---------
function validateSurveyTask(p_oSurveyTask, p_sSPC) {

	var szMsg = "";
	
	//mandatory fields
    var sMandatoriesMsg = "";
    var aMandatories = getMandatoryFieldsForSurveyTasks();
	sMandatoriesMsg = validateMandatoryObjectAttributes(p_oSurveyTask, aMandatories, p_sSPC);
        
    if (!sMandatoriesMsg.equals("")){
        szMsg = sMandatoriesMsg;
    }
	
	//check if end date is before start date
    var sStartEndDateMessage = validateStartEndDate(p_oSurveyTask, Constants.AT_SURVEYTASK_START_DATE, Constants.AT_SURVEYTASK_END_DATE, p_sSPC);
	if (!sStartEndDateMessage.equals("")) {
        if (szMsg != "") {szMsg += "\r\n" + "\r\n";}
        szMsg += sStartEndDateMessage;
    }
	
	//linked objects
	var sLinkedQuestionnaireTemplatesMsg = "";
    var aConnectedQuestionnaireTemplates = new Array();
    if (g_surveyTask2questionnaireTemplateHashMap.get(p_oSurveyTask) != null) {aConnectedQuestionnaireTemplates = g_surveyTask2questionnaireTemplateHashMap.get(p_oSurveyTask);}
    sLinkedQuestionnaireTemplatesMsg = getConnectionCountValidationOutput(SM_OCCURRENCE_MIN_SURVEYTASK_TO_QUESTIONNAIRE_TEMPLATE, SM_OCCURRENCE_MAX_SURVEYTASK_TO_QUESTIONNAIRE_TEMPLATE, p_oSurveyTask, aConnectedQuestionnaireTemplates, getString("TEXT_19"), SPC1);      
    if (!sLinkedQuestionnaireTemplatesMsg.equals("")) {
        if (szMsg != "") {szMsg += "\r\n" + "\r\n";}
        szMsg += sLinkedQuestionnaireTemplatesMsg;
    }
	
    var sLinkedQuestionnaireOwnersMsg = "";
    var aConnectedQuestionnaireOwners = new Array();
    if (g_surveyTask2questionnaireownerHashMap.get(p_oSurveyTask) != null) {aConnectedQuestionnaireOwners = g_surveyTask2questionnaireownerHashMap.get(p_oSurveyTask);}
    sLinkedQuestionnaireOwnersMsg = getConnectionCountValidationOutput(SM_OCCURRENCE_MIN_SURVEYTASK_TO_QUESTIONNAIREOWNER, SM_OCCURRENCE_MAX_SURVEYTASK_TO_QUESTIONNAIREOWNER, p_oSurveyTask, aConnectedQuestionnaireOwners, getString("TEXT_20"), SPC1);      
    if (!sLinkedQuestionnaireOwnersMsg.equals("")) {
        if (szMsg != "") {szMsg += "\r\n" + "\r\n";}
        szMsg += sLinkedQuestionnaireOwnersMsg;
    }

    var sLinkedSurveyReviewersMsg = "";
    var aConnectedSurveyReviewers = new Array();
    if (g_surveyTask2surveyreviewerHashMap.get(p_oSurveyTask) != null) {aConnectedSurveyReviewers = g_surveyTask2surveyreviewerHashMap.get(p_oSurveyTask);}
    sLinkedSurveyReviewersMsg = getConnectionCountValidationOutput(SM_OCCURRENCE_MIN_SURVEYTASK_TO_SURVEYREVIEWER, SM_OCCURRENCE_MAX_SURVEYTASK_TO_SURVEYREVIEWER, p_oSurveyTask, aConnectedSurveyReviewers, getString("TEXT_21"), SPC1);      
    if (!sLinkedSurveyReviewersMsg.equals("")) {
        if (szMsg != "") {szMsg += "\r\n" + "\r\n";}
        szMsg += sLinkedSurveyReviewersMsg;
    }
	
	var sLinkedSurveyManagerMsg = "";
    var aConnectedSurveyManagers = new Array();
	if (g_surveyTask2surveymanagerHashMap.get(p_oSurveyTask) != null) {aConnectedSurveyManagers = g_surveyTask2surveymanagerHashMap.get(p_oSurveyTask);}
    sLinkedSurveyManagerMsg = getConnectionCountValidationOutput(SM_OCCURRENCE_MIN_SURVEYTASK_TO_SURVEYMANAGER, SM_OCCURRENCE_MAX_SURVEYTASK_TO_SURVEYMANAGER, p_oSurveyTask, aConnectedSurveyManagers, getString("TEXT_31"), p_sSPC);
	if (!sLinkedSurveyManagerMsg.equals("")) {
        if (szMsg != "") {szMsg += "\r\n" + "\r\n";}
        szMsg += sLinkedSurveyManagerMsg;
    }
     
	var sLinkedRisksMsg = "";
    var aConnectedRisks = new Array();
    if (g_surveyTask2riskHashMap.get(p_oSurveyTask) != null) {aConnectedRisks = g_surveyTask2riskHashMap.get(p_oSurveyTask);}
    sLinkedRisksMsg = getConnectionCountValidationOutput(SM_OCCURRENCE_MIN_SURVEYTASK_TO_RISK, SM_OCCURRENCE_MAX_SURVEYTASK_TO_RISK, p_oSurveyTask, aConnectedRisks, getString("TEXT_22"), SPC1);      
    if (!sLinkedRisksMsg.equals("")) {
        if (szMsg != "") {szMsg += "\r\n" + "\r\n";}
        szMsg += sLinkedRisksMsg;
    }
	
	var sLinkedControlsMsg = "";
    var aConnectedControls = new Array();
    if (g_surveyTask2controlHashMap.get(p_oSurveyTask) != null) {aConnectedControls = g_surveyTask2controlHashMap.get(p_oSurveyTask);}
    sLinkedControlsMsg = getConnectionCountValidationOutput(SM_OCCURRENCE_MIN_SURVEYTASK_TO_CONTROL, SM_OCCURRENCE_MAX_SURVEYTASK_TO_CONTROL, p_oSurveyTask, aConnectedControls, getString("TEXT_23"), SPC1);      
    if (!sLinkedControlsMsg.equals("")) {
        if (szMsg != "") {szMsg += "\r\n" + "\r\n";}
        szMsg += sLinkedControlsMsg;
    }
	
	var sLinkedTestdefinitionsMsg = "";
    var aConnectedTestdefinitions = new Array();
    if (g_surveyTask2testdefinitionHashMap.get(p_oSurveyTask) != null) {aConnectedTestdefinitions = g_surveyTask2testdefinitionHashMap.get(p_oSurveyTask);}
    sLinkedTestdefinitionsMsg = getConnectionCountValidationOutput(SM_OCCURRENCE_MIN_SURVEYTASK_TO_TESTDEFINITION, SM_OCCURRENCE_MAX_SURVEYTASK_TO_TESTDEFINITION, p_oSurveyTask, aConnectedTestdefinitions, getString("TEXT_24"), SPC1);      
    if (!sLinkedTestdefinitionsMsg.equals("")) {
        if (szMsg != "") {szMsg += "\r\n" + "\r\n";}
        szMsg += sLinkedTestdefinitionsMsg;
    }
	
	var sLinkedAppSysTypesMsg = "";
    var aConnectedAppSysTypes = new Array();
    if (g_surveyTask2appSysTypeHashMap.get(p_oSurveyTask) != null) {aConnectedAppSysTypes = g_surveyTask2appSysTypeHashMap.get(p_oSurveyTask);}
    sLinkedAppSysTypesMsg = getConnectionCountValidationOutput(SM_OCCURRENCE_MIN_SURVEYTASK_TO_APPSYS, SM_OCCURRENCE_MAX_SURVEYTASK_TO_APPSYS, p_oSurveyTask, aConnectedAppSysTypes, getString("TEXT_25"), SPC1);      
    if (!sLinkedAppSysTypesMsg.equals("")) {
        if (szMsg != "") {szMsg += "\r\n" + "\r\n";}
        szMsg += sLinkedAppSysTypesMsg;
    }
	
	var sLinkedRegulationsMsg = "";
    var aConnectedRegulations = new Array();
    if (g_surveyTask2regulationHashMap.get(p_oSurveyTask) != null) {aConnectedRegulations = g_surveyTask2regulationHashMap.get(p_oSurveyTask);}
    sLinkedRegulationsMsg = getConnectionCountValidationOutput(SM_OCCURRENCE_MIN_SURVEYTASK_TO_REGULATION, SM_OCCURRENCE_MAX_SURVEYTASK_TO_REGULATION, p_oSurveyTask, aConnectedRegulations, getString("TEXT_26"), SPC1);      
    if (!sLinkedRegulationsMsg.equals("")) {
        if (szMsg != "") {szMsg += "\r\n" + "\r\n";}
        szMsg += sLinkedRegulationsMsg;
    }
	
	var sLinkedOrgUnitsMsg = "";
    var aConnectedOrgUnits = new Array();
    if (g_surveyTask2orgUnitHashMap.get(p_oSurveyTask) != null) {aConnectedOrgUnits = g_surveyTask2orgUnitHashMap.get(p_oSurveyTask);}
    sLinkedOrgUnitsMsg = getConnectionCountValidationOutput(SM_OCCURRENCE_MIN_SURVEYTASK_TO_ORGANIZATION, SM_OCCURRENCE_MAX_SURVEYTASK_TO_ORGANIZATION, p_oSurveyTask, aConnectedOrgUnits, getString("TEXT_27"), SPC1);      
    if (!sLinkedOrgUnitsMsg.equals("")) {
        if (szMsg != "") {szMsg += "\r\n" + "\r\n";}
        szMsg += sLinkedOrgUnitsMsg;
    }
	
	var sLinkedProcessesMsg = "";
    var aConnectedProcesses = new Array();
    if (g_surveyTask2processHashMap.get(p_oSurveyTask) != null) {aConnectedProcesses = g_surveyTask2processHashMap.get(p_oSurveyTask);}
    sLinkedProcessesMsg = getConnectionCountValidationOutput(SM_OCCURRENCE_MIN_SURVEYTASK_TO_PROCESS, SM_OCCURRENCE_MAX_SURVEYTASK_TO_PROCESS, p_oSurveyTask, aConnectedProcesses, getString("TEXT_28"), SPC1);      
    if (!sLinkedProcessesMsg.equals("")) {
        if (szMsg != "") {szMsg += "\r\n" + "\r\n";}
        szMsg += sLinkedProcessesMsg;
    }
	
	var sLinkedRiskCategoriesMsg = "";
    var aConnectedRiskCategories = new Array();
    if (g_surveyTask2riskCategoryHashMap.get(p_oSurveyTask) != null) {aConnectedRiskCategories = g_surveyTask2riskCategoryHashMap.get(p_oSurveyTask);}
    sLinkedRiskCategoriesMsg = getConnectionCountValidationOutput(SM_OCCURRENCE_MIN_SURVEYTASK_TO_RISKCAT, SM_OCCURRENCE_MAX_SURVEYTASK_TO_RISKCAT, p_oSurveyTask, aConnectedRiskCategories, getString("TEXT_29"), SPC1);      
    if (!sLinkedRiskCategoriesMsg.equals("")) {
        if (szMsg != "") {szMsg += "\r\n" + "\r\n";}
        szMsg += sLinkedRiskCategoriesMsg;
    }
	
	
	//in case of errors add header with the survey task name
    if (szMsg.length > 0) {
        var sOptionHeader = getString("TEXT_30") + " \"" + p_oSurveyTask.Name(g_nLoc) + "\":";
        szMsg = p_sSPC + sOptionHeader + "\r\n"
                + getUnderline(sOptionHeader, p_sSPC) + "\r\n"
                + szMsg + "\r\n";
    }
	
	return szMsg; 
}