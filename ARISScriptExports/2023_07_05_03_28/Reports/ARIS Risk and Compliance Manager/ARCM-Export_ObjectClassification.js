/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

/*----------------------------------Global sets----------------------------------------
---------------------------------------------------------------------------------------*/
const PROCESS_MODEL = "PROCESS_MODEL";
const ORG_UNIT_MODEL = "ORG_UNIT_MODEL";
const REGULATION_MODEL = "REGULATION_MODEL";
const DATA_CLUSTER_MODEL = "DATA_CLUSTER_MODEL";
const RISK_CATEGORY_MODEL = "RISK_CATEGORY_MODEL";
const APPLICATION_SYSTEM_TYPE_MODEL = "APPLICATION_SYSTEM_TYPE_MODEL";


var roleMappingXMLElement = null;

//global set to determine if a link between two ObjDefs is deactivated
//in order to speed up things this set is filled with the concatenated GUIDs if the link between them is deactivated
var g_deactivatedLinkSet = new java.util.HashSet();

var g_userGroupSet      = new java.util.HashSet();     // global set of user groups
var g_userSet           = new java.util.HashSet();     // global set of users

// global GDPR stuff
var g_processGDPRSet    = new java.util.HashSet();     // global set of processes for which the GDPR classification was already done 
var g_dataClusterSet    = new java.util.HashSet();     // global set of data clusters

// Hierarchy parent to child relation
var g_hierarchyParent2hierarchyChildrenHashMap = new java.util.HashMap(); //global HashMap for linkage parent hierarchy element->child hierarchy elements - Format: ObjDef | HashSet <ObjDef>
// Hierarchy child to parent relation
var g_hierarchyChild2hierarchyParentsHashMap = new java.util.HashMap(); //global HashMap for linkage child hierarchy element->parent hierarchy elements - Format: ObjDef | HashSet <ObjDef>

// Hierarchy usergroup assignments
var g_hierarchy2TestAuditorHashMap  = new java.util.HashMap(); //global HashMap for linkage hierarchy element->testauditor usergroups - Format: ObjDef | JSArray <ObjDef>
var g_hierarchy2HierarchyOwnerHashMap  = new java.util.HashMap(); //global HashMap for linkage hierarchy element->hierarchy owner usergroups - Format: ObjDef | JSArray <ObjDef>

var g_allowedModels = {}; // global map for allowed models for classification. Build from aris2arcm-mapping.xml (allowedModels element in xml). key: PROCESS_MODEL => value: set of model type number
var g_allowedObjects = {}; // global map for allowed objects for classification. Build from aris2arcm-mapping.xml (allowedModels element in xml). key: PROCESS_MODEL => value: set of object type number
var g_allowedSymbols = {}; // global map for allowed symbols for classification. Build from aris2arcm-mapping.xml (allowedModels element in xml). key: PROCESS_MODEL => value: set of symbol type number

var g_processModelClassified = new java.util.HashSet(); //global HashMap for process models that are classified.
var g_regulationModelClassified = new java.util.HashSet(); //global HashMap for regulation models that are classified.
var g_riskCategoryModelClassified = new java.util.HashSet(); //global HashMap for risk category models that are classified.
var g_orgUnitModelClassified = new java.util.HashSet(); //global HashMap for org unit models that are classified.
var g_applicationSystemTypeModelClassified = new java.util.HashSet(); //global HashMap for application system type models that are classified.
var g_dataClusterModelClassified = new java.util.HashSet(); //global HashMap for data cluster type models that are classified.

//---------- I. Test management ----------
// Test management HashSets
var g_riskSet                        = new java.util.HashSet();     // global set of risks
var g_controlSet                     = new java.util.HashSet();     // global set of controls
var g_controlExecutionTaskSet        = new java.util.HashSet();     // global set of control execution tasks
var g_testdefinitionSet              = new java.util.HashSet();     // global set of testdefinitions

// risk HashSet for preventing multiple processing of risks for Test management
var g_processedRisksTestMode = new java.util.HashSet(); // global set of already processed risks by test management mode
// control HashSets for preventing multiple processing of controls for Test management
var g_processedControlsTestMode = new java.util.HashSet(); // global set of already processed controls by test management mode


// sets for erroneous objects
var g_error_risk_genericSpecific = new java.util.HashSet(); //global set of risks that are modeled both as generic and specific

// Test management hierarchies
var g_testerSet         = new java.util.HashSet();     // global set of testers
var g_regulationSet     = new java.util.HashSet();     // global set of regulations
var g_orgUnitSet        = new java.util.HashSet();     // global set of organisational units
var g_processSet        = new java.util.HashSet();     // global set of processes

// Test management assignment HashSets
var g_risk2controlsHashMap              = new java.util.HashMap(); //global HashMap for linkage risk->controls - Format: ObjDef | JSArray <ObjDef>

var g_control2processesHashMap          = new java.util.HashMap(); //global HashMap for linkage control->processes - Format: ObjDef | JSArray <ObjDef>
var g_control2regulationsHashMap        = new java.util.HashMap(); //global HashMap for linkage control->regulations - Format: ObjDef | JSArray <ObjDef>
var g_risk2processesHashMap             = new java.util.HashMap(); //global HashMap for linkage risk->processes - Format: ObjDef | JSArray <ObjDef>
var g_risk2riskManagerHashMap           = new java.util.HashMap(); //global HashMap for linkage risk->risk manager groups - Format: ObjDef | JSArray <ObjDef>
var g_risk2regulationsHashMap           = new java.util.HashMap(); //global HashMap for linkage risk->regulations - Format: ObjDef | JSArray <ObjDef>
var g_risk2orgunitHashMap               = new java.util.HashMap(); //global HashMap for linkage risk->orgunits - Format: ObjDef | JSArray <ObjDef>

var g_specificRisk2genericRiskHashMap   = new java.util.HashMap(); //global HashMap for linkage specific risk->generic risk - Format: ObjDef | ObjDef

var g_control2risksHashMap              = new java.util.HashMap(); //global HashMap for linkage control->risks - Format: ObjDef | JSArray <ObjDef>
var g_control2testdefinitionsHashMap    = new java.util.HashMap(); //global HashMap for linkage control->testdefinitions - Format: ObjDef | JSArray <ObjDef>
var g_control2controlManagerHashMap     = new java.util.HashMap(); //global HashMap for linkage control->control manager groups - Format: ObjDef | JSArray <ObjDef>
var g_control2controlExecutionTasksHashMap     = new java.util.HashMap(); //global HashMap for linkage control->control execution tasks - Format: ObjDef | JSArray <ObjDef>

var g_controlExecutionTask2controlHashMap               = new java.util.HashMap(); //global HashMap for linkage control execution task->control - Format: ObjDef | JSArray <ObjDef>
var g_controlExecutionTask2controlExecutionOwnerHashMap = new java.util.HashMap(); //global HashMap for linkage control execution task->control execution owner - Format: ObjDef | JSArray <ObjDef>
var g_controlExecutionTask2orgunitHashMap               = new java.util.HashMap(); //global HashMap for linkage control execution task->orgunit - Format: ObjDef | JSArray <ObjDef>

var g_testdefinition2orgunitHashMap     = new java.util.HashMap(); //global HashMap for linkage testdefinition->orgunits - Format: ObjDef | JSArray <ObjDef>
var g_testdefinition2controlsHashMap    = new java.util.HashMap(); //global HashMap for linkage testdefinition->controls - Format: ObjDef | JSArray <ObjDef>
var g_testdefinition2testManagerHashMap = new java.util.HashMap(); //global HashMap for linkage testdefinition->usergroups - Format: ObjDef | JSArray <ObjDef>
var g_testdefinition2testerHashMap      = new java.util.HashMap(); //global HashMap for linkage testdefinition->usergroups - Format: ObjDef | JSArray <ObjDef>
var g_testdefinition2testReviewerHashMap  = new java.util.HashMap(); //global HashMap for linkage testdefinition->usergroups - Format: ObjDef | JSArray <ObjDef>

// Test management hierarchy usergroup assignments
var g_hierarchy2SignoffOwnerHashMap         = new java.util.HashMap(); //global HashMap for linkage hierarchy element->signoff owner usergroups - Format: ObjDef | JSArray <ObjDef>
var g_hierarchy2TesterGroupHashMap          = new java.util.HashMap(); //global HashMap for linkage hierarchy element->tester usergroups - Format: ObjDef | JSArray <ObjDef>
var g_testerGroup2TesterHierarchyHashMap    = new java.util.HashMap(); //global HashMap for linkage tester usergroups -> tester hierarchy element - Format: ObjDef | JSArray <ObjDef> 


//---------- II. Risk management ----------
// risk HashSets for preventing multiple processing of risks for Risk management
var g_processedRisksRiskMode = new java.util.HashSet(); // global set of already processed risks by risk management mode
// control HashSets for preventing multiple processing of controls
var g_processedControlsRiskMode = new java.util.HashSet(); // global set of already processed controls by risk management mode

// Risk management hierarchies
var g_appSysSet         = new java.util.HashSet();     // global set of application system types
var g_riskCategorySet   = new java.util.HashSet();     // global set of risk categories

// Risk management assignment HashSets
var g_risk2appSysTypeHashMap            = new java.util.HashMap(); //global HashMap for linkage risk->appsystypes - Format: ObjDef | JSArray <ObjDef>
var g_risk2riskCategoryHashMap          = new java.util.HashMap(); //global HashMap for linkage risk -> risk category - Format: ObjDef | JSArray <ObjDef>
var g_risk2riskOwnerHashMap             = new java.util.HashMap(); //global HashMap for linkage risk->usergroups - Format: ObjDef | JSArray <ObjDef>
var g_risk2riskReviewerHashMap          = new java.util.HashMap(); //global HashMap for linkage risk->usergroups - Format: ObjDef | JSArray <ObjDef>


//---------- III. Survey management ----------
var g_questionnaireTemplateSet  = new java.util.HashSet();     // global set of questionnaire templates
var g_sectionSet                = new java.util.HashSet();     // global set of sections
var g_questionSet               = new java.util.HashSet();     // global set of question
var g_optionSetSet              = new java.util.HashSet();     // global set of option sets
var g_optionSet                 = new java.util.HashSet();     // global set of options
var g_surveyTaskSet             = new java.util.HashSet();     // global set of survey tasks

var g_questionnaireTemplate2sectionsHashMap         = new java.util.HashMap(); //global HashMap for linkage questionnaire template->section - Format: ObjDef | JSArray <ObjDef>
var g_questionnaireTemplate2surveymanagerHashMap    = new java.util.HashMap(); //global HashMap for linkage questionnaire template->survey manager - Format: ObjDef | JSArray <ObjDef>
var g_questionnaireTemplate2sectionsHashMap         = new java.util.HashMap(); //global HashMap for linkage questionnaire template->section - Format: ObjDef | JSArray <ObjDef>
var g_section2questionnaireTemplateHashMap          = new java.util.HashMap(); //global HashMap for linkage section->questionnaire templates - Format: ObjDef | JSArray <ObjDef>
var g_section2subsectionsHashMap                    = new java.util.HashMap(); //global HashMap for linkage section->sub sections - Format: ObjDef | JSArray <ObjDef>
var g_section2topsectionsHashMap                    = new java.util.HashMap(); //global HashMap for linkage section->top sections - Format: ObjDef | JSArray <ObjDef>
var g_section2questionsHashMap                      = new java.util.HashMap(); //global HashMap for linkage section->questions - Format: ObjDef | JSArray <ObjDef>
var g_question2sectionsHashMap                      = new java.util.HashMap(); //global HashMap for linkage question->sections - Format: ObjDef | JSArray <ObjDef>
var g_question2optionsHashMap                       = new java.util.HashMap(); //global HashMap for linkage question->options - Format: ObjDef | JSArray <ObjDef>
var g_question2optionSetsHashMap                    = new java.util.HashMap(); //global HashMap for linkage question->optionSets - Format: ObjDef | JSArray <ObjDef>
var g_optionSet2optionsHashMap                      = new java.util.HashMap(); //global HashMap for linkage optionSet->options - Format: ObjDef | JSArray <ObjDef>

var g_questionnaireTemplate2surveyTaskHashMap	= new java.util.HashMap(); //global HashMap for linkage questionnaire template->survey task - Format: ObjDef | JSArray <ObjDef> 
var g_surveyTask2questionnaireTemplateHashMap	= new java.util.HashMap(); //global HashMap for linkage survey task->questionnaire template - Format: ObjDef | JSArray <ObjDef> 
var g_surveyTask2questionnaireownerHashMap		= new java.util.HashMap(); //global HashMap for linkage survey task->questionnaire owner - Format: ObjDef | JSArray <ObjDef> 
var g_surveyTask2surveyreviewerHashMap			= new java.util.HashMap(); //global HashMap for linkage survey task->survey reviewer - Format: ObjDef | JSArray <ObjDef> 
var g_surveyTask2surveymanagerHashMap           = new java.util.HashMap(); //global HashMap for linkage survey task->survey manager - Format: ObjDef | JSArray <ObjDef> 
var g_surveyTask2riskHashMap					= new java.util.HashMap(); //global HashMap for linkage survey task->survey reviewer - Format: ObjDef | JSArray <ObjDef> 
var g_surveyTask2controlHashMap					= new java.util.HashMap(); //global HashMap for linkage survey task->survey reviewer - Format: ObjDef | JSArray <ObjDef> 
var g_surveyTask2testdefinitionHashMap			= new java.util.HashMap(); //global HashMap for linkage survey task->survey reviewer - Format: ObjDef | JSArray <ObjDef> 
var g_surveyTask2regulationHashMap				= new java.util.HashMap(); //global HashMap for linkage survey task->regulations - Format: ObjDef | JSArray <ObjDef>
var g_surveyTask2orgUnitHashMap          		= new java.util.HashMap(); //global HashMap for linkage survey task->orgunit - Format: ObjDef | JSArray <ObjDef>
var g_surveyTask2processHashMap          		= new java.util.HashMap(); //global HashMap for linkage survey task->processes - Format: ObjDef | JSArray <ObjDef>
var g_surveyTask2appSysTypeHashMap       		= new java.util.HashMap(); //global HashMap for linkage survey task->appsystypes - Format: ObjDef | JSArray <ObjDef>
var g_surveyTask2riskCategoryHashMap     		= new java.util.HashMap(); //global HashMap for linkage survey task->risk category - Format: ObjDef | JSArray <ObjDef>

//recursion stop sets fur survey management
//needed because of the re-usage of objects within different questionnaire templates and the cyclic dependencies between activating and depending questions
//not part of store/restore
var g_sectionRecursionStop                      = new java.util.HashSet(); //notes which section was already processed in which questionnaire template context - Format: <QT GUID> + "|" + <Section GUID>
var g_questionRecursionStop                     = new java.util.HashSet(); //notes which question was already processed in which questionnaire template context - Format: <QT GUID> + "|" + <Question GUID>
var g_optionSetRecursionStop                    = new java.util.HashSet(); //notes which option set was already processed in which questionnaire template context - Format: <QT GUID> + "|" + <Option Set GUID>
var g_optionRecursionStop                       = new java.util.HashSet(); //notes which option was already processed in which questionnaire template context - Format: <QT GUID> + "|" + <Option GUID>

var g_explicitPositionSections                  = new java.util.HashMap(); //notes which sections have an explicit position inside which questionnaire template - Format: ObjDef | JSArray <ObjDef>
var g_explicitPositionQuestions                 = new java.util.HashMap(); //notes which questions have an explicit position inside which questionnaire template - Format: ObjDef | JSArray <ObjDef>

var g_possibleImplicitTopLevelSections          = new java.util.HashMap() //notes which top level sections are possibly implictly placed in which questionnaire template - Format: ObjDef | JSArray <JS objects 'implicitPlacing'>
var g_possibleImplicitSections                  = new java.util.HashMap() //notes which sections are possibly implictly placed in which questionnaire template - Format: ObjDef | JSArray <JS objects 'implicitPlacing'>
var g_possibleImplicitQuestions                 = new java.util.HashMap() //notes which questions are possibly implictly placed in which questionnaire template - Format: ObjDef | JSArray <JS objects 'implicitPlacing'>

var g_realImplicitSections                      = new java.util.HashMap() //notes which sections are really implictly placed in which questionnaire template - Format: ObjDef | JSArray <JS objects 'implicitPlacing'>
var g_realImplicitQuestions                     = new java.util.HashMap() //notes which questions are really implictly placed in which questionnaire template - Format: ObjDef | JSArray <JS objects 'implicitPlacing'>

//dependent visibility
var g_question2VisContextQuestionsHashMap       = new java.util.HashMap(); //global HashMap for linkage question->visibility context questions - Format: ObjDef | JSArray <ObjDef>
var g_section2VisContextQuestionsHashMap        = new java.util.HashMap(); //global HashMap for linkage section->visibility context questions - Format: ObjDef | JSArray <ObjDef>

//---------- IV. Audit management ----------
var g_auditTemplateSet 			= new java.util.HashSet(); // global set of audit templates
var g_auditStepTemplateSet 		= new java.util.HashSet(); // global set of audit step templates

var g_auditTemplate2AuditStepTemplateHashMap		= new java.util.HashMap(); //global HashMap for linkage audit templates->audit step templates - Format: ObjDef | JSArray <ObjDef>
var g_auditStepTemplate2AuditTemplateHashMap		= new java.util.HashMap(); //global HashMap for linkage audit step templates->audit templates - Format: ObjDef | JSArray <ObjDef>
var g_auditStepTemplate2subAuditStepTemplateHashMap	= new java.util.HashMap(); //global HashMap for linkage audit step templates->sub audit templates - Format: ObjDef | JSArray <ObjDef>
var g_auditStepTemplate2topAuditStepTemplateHashMap	= new java.util.HashMap(); //global HashMap for linkage audit step templates->top audit templates - Format: ObjDef | JSArray <ObjDef>  
var g_auditTemplate2auditOwnerHashMap 				= new java.util.HashMap(); //global HashMap for linkage audit templates->usergroups - Format: ObjDef | JSArray <ObjDef> 
var g_auditTemplate2auditReviewerHashMap 			= new java.util.HashMap(); //global HashMap for linkage audit templates->usergroups - Format: ObjDef | JSArray <ObjDef> 
var g_auditTemplate2auditAuditorHashMap 			= new java.util.HashMap(); //global HashMap for linkage audit templates->usergroups - Format: ObjDef | JSArray <ObjDef>
var g_auditStepTemplate2auditStepOwnerHashMap 		= new java.util.HashMap(); //global HashMap for linkage audit templates->usergroups - Format: ObjDef | JSArray <ObjDef>  

//scope assignments
var g_auditTemplate2regulationHashMap		= new java.util.HashMap(); //global HashMap for linkage audit templates->regulations - Format: ObjDef | JSArray <ObjDef>
var g_auditTemplate2orgUnitHashMap          = new java.util.HashMap(); //global HashMap for linkage audit templates->orgunit - Format: ObjDef | JSArray <ObjDef>
var g_auditTemplate2processHashMap          = new java.util.HashMap(); //global HashMap for linkage audit templates->processes - Format: ObjDef | JSArray <ObjDef>
var g_auditTemplate2appSysTypeHashMap       = new java.util.HashMap(); //global HashMap for linkage audit templates->appsystypes - Format: ObjDef | JSArray <ObjDef>
var g_auditTemplate2riskCategoryHashMap     = new java.util.HashMap(); //global HashMap for linkage audit templates->risk category - Format: ObjDef | JSArray <ObjDef>

var g_auditStepTemplate2regulationHashMap   = new java.util.HashMap(); //global HashMap for linkage audit step templates->regulations - Format: ObjDef | JSArray <ObjDef>
var g_auditStepTemplate2orgUnitHashMap      = new java.util.HashMap(); //global HashMap for linkage audit step templates->orgunit - Format: ObjDef | JSArray <ObjDef>
var g_auditStepTemplate2processHashMap      = new java.util.HashMap(); //global HashMap for linkage audit step templates->processes - Format: ObjDef | JSArray <ObjDef>
var g_auditStepTemplate2appSysTypeHashMap   = new java.util.HashMap(); //global HashMap for linkage audit step templates->appsystypes - Format: ObjDef | JSArray <ObjDef>
var g_auditStepTemplate2riskCategoryHashMap = new java.util.HashMap(); //global HashMap for linkage audit step templates->risk category - Format: ObjDef | JSArray <ObjDef>


//---------- V. Policy management ----------
var g_policyDefinitionSet 			       = new java.util.HashSet(); // global set of policy definitions

var g_policyDefinition2policyOwnerHashMap       = new java.util.HashMap(); //global HashMap for linkage policy definitions->usergroups - Format: ObjDef | JSArray <ObjDef> 
var g_policyDefinition2policyApproverHashMap    = new java.util.HashMap(); //global HashMap for linkage policy definitions->usergroups - Format: ObjDef | JSArray <ObjDef> 
var g_policyDefinition2policyAddresseeHashMap   = new java.util.HashMap(); //global HashMap for linkage policy definitions->usergroups - Format: ObjDef | JSArray <ObjDef> 
var g_policyDefinition2policyReviewerHashMap    = new java.util.HashMap(); //global HashMap for linkage policy definitions->usergroups - Format: ObjDef | JSArray <ObjDef> 
var g_policyDefinition2policyAuditorHashMap     = new java.util.HashMap(); //global HashMap for linkage policy definitions->usergroups - Format: ObjDef | JSArray <ObjDef> 

var g_policyDefinition2riskHashMap         = new java.util.HashMap(); //global HashMap for linkage policy definitions->risks - Format: ObjDef | JSArray <ObjDef>
var g_policyDefinition2regulationHashMap   = new java.util.HashMap(); //global HashMap for linkage policy definitions->regulations - Format: ObjDef | JSArray <ObjDef>
var g_policyDefinition2orgUnitHashMap      = new java.util.HashMap(); //global HashMap for linkage policy definitions->orgunit - Format: ObjDef | JSArray <ObjDef>
var g_policyDefinition2processHashMap      = new java.util.HashMap(); //global HashMap for linkage policy definitions->processes - Format: ObjDef | JSArray <ObjDef>
var g_policyDefinition2appSysTypeHashMap   = new java.util.HashMap(); //global HashMap for linkage policy definitions->appsystypes - Format: ObjDef | JSArray <ObjDef>



var g_testManagementHashSetsFilled      = false;
var g_riskManagementHashSetsFilled      = false;
var g_surveyManagementHashSetsFilled    = false;
var g_auditManagementHashSetsFilled     = false;
var g_policyManagementHashSetsFilled    = false;
var g_gdprHashSetsFilled                = false;

/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                        ~~~~~~~ Data structures ~~~~~~~
                        ~~~~~~~ Survey management ~~~~~~~
  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
function implicitPlacing(p_oObjDefToPlace, p_oPositioningSection, p_oActivatingSection, p_oActivatingQuestion, p_aContextQuestions) {
    this.oObjDefToPlace = p_oObjDefToPlace;
    this.oPositioningSection = p_oPositioningSection;
    this.oActivatingSection = p_oActivatingSection;
    this.oActivatingQuestion = p_oActivatingQuestion;
    this.aContextQuestions = p_aContextQuestions;
    return this;
}
implicitPlacing.prototype.toString = function impPlacToString() {
  var sObjDefToPlace = "<none>";
  if (this.oObjDefToPlace != null) {sObjDefToPlace = this.oObjDefToPlace.Name(g_nLoc);}
  var sPositioningSection = "<none>";
  if (this.oPositioningSection != null) {sPositioningSection = this.oPositioningSection.Name(g_nLoc);}
  var sActivatingSection = "<none>";
  if (this.oActivatingSection != null) {sActivatingSection = this.oActivatingSection.Name(g_nLoc);}
  var sActivatingQuestion = "<none>";
  if (this.oActivatingQuestion != null) {sActivatingQuestion = this.oActivatingQuestion.Name(g_nLoc);}
  var sContextQuestions = "<none>";
  if (this.aContextQuestions != null && this.aContextQuestions.length > 0) {
      sContextQuestions = "[";
      for (var i=0; i<this.aContextQuestions.length; i++) {
          sContextQuestions += this.aContextQuestions[i].Name(g_nLoc);
          if (i<this.aContextQuestions.length - 1) {sContextQuestions += ", ";}
      }
      sContextQuestions += "]";
  }
  var ret = 'ObjDefToPlace: ' + sObjDefToPlace + ' ---- Position: ' + sPositioningSection + ' ---- Act.Section: ' + sActivatingSection + " ---- Act.Question: " + sActivatingQuestion + ' ---- Context: ' + sContextQuestions;
  return ret;
}

/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ~~~~~~~ Restore HashSets and HashMaps from report context if possible ~~~~~~~
                        ~~~~~~~ Test management ~~~~~~~
  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
/**
  * Restore all test management HashSets and HashMaps
  */
function restoreTestManagementFromReportContext() {
    
	if ( Context.getProperty("g_testManagementHashSetsFilled") == null 
         || !Context.getProperty("g_testManagementHashSetsFilled").equals("true") ) {
         return false;
    }
    
    g_deactivatedLinkSet = restoreStringHashSetFromProperty("g_deactivatedLinkSet", ArisData.getActiveDatabase());
    g_userGroupSet = restoreHashSetFromProperty("g_userGroupSet", ArisData.getActiveDatabase());
    g_userSet = restoreHashSetFromProperty("g_userSet", ArisData.getActiveDatabase());
    g_riskSet = restoreHashSetFromProperty("g_riskSet", ArisData.getActiveDatabase());
    g_controlSet = restoreHashSetFromProperty("g_controlSet", ArisData.getActiveDatabase());
    g_controlExecutionTaskSet = restoreHashSetFromProperty("g_controlExecutionTaskSet", ArisData.getActiveDatabase()); 
    g_testdefinitionSet = restoreHashSetFromProperty("g_testdefinitionSet", ArisData.getActiveDatabase());    
    g_processedRisksTestMode = restoreHashSetFromProperty("g_processedRisksTestMode", ArisData.getActiveDatabase());
    g_processedControlsTestMode = restoreHashSetFromProperty("g_processedControlsTestMode", ArisData.getActiveDatabase());
	g_error_risk_genericSpecific = restoreHashSetFromProperty("g_error_risk_genericSpecific", ArisData.getActiveDatabase());
    g_testerSet = restoreHashSetFromProperty("g_testerSet", ArisData.getActiveDatabase());
    g_regulationSet = restoreHashSetFromProperty("g_regulationSet", ArisData.getActiveDatabase());
    g_orgUnitSet = restoreHashSetFromProperty("g_orgUnitSet", ArisData.getActiveDatabase());
    g_processSet = restoreHashSetFromProperty("g_processSet", ArisData.getActiveDatabase());
    
    g_risk2controlsHashMap = restoreHashMapFromProperty("g_risk2controlsHashMap", ArisData.getActiveDatabase());
    g_risk2processesHashMap = restoreHashMapFromProperty("g_risk2processesHashMap", ArisData.getActiveDatabase());
    g_risk2riskManagerHashMap = restoreHashMapFromProperty("g_risk2riskManagerHashMap", ArisData.getActiveDatabase());
    g_risk2regulationsHashMap = restoreHashMapFromProperty("g_risk2regulationsHashMap", ArisData.getActiveDatabase());
    g_risk2orgunitHashMap = restoreHashMapFromProperty("g_risk2orgunitHashMap", ArisData.getActiveDatabase());
    g_specificRisk2genericRiskHashMap = restoreHashMapFromProperty("g_specificRisk2genericRiskHashMap", ArisData.getActiveDatabase());
    g_control2processesHashMap = restoreHashMapFromProperty("g_control2processesHashMap", ArisData.getActiveDatabase());
	g_control2regulationsHashMap = restoreHashMapFromProperty("g_control2regulationsHashMap", ArisData.getActiveDatabase());
	g_control2risksHashMap = restoreHashMapFromProperty("g_control2risksHashMap", ArisData.getActiveDatabase());
    g_control2testdefinitionsHashMap = restoreHashMapFromProperty("g_control2testdefinitionsHashMap", ArisData.getActiveDatabase());
    g_control2controlManagerHashMap = restoreHashMapFromProperty("g_control2controlManagerHashMap", ArisData.getActiveDatabase());   
    g_control2controlExecutionTasksHashMap = restoreHashMapFromProperty("g_control2controlExecutionTasksHashMap", ArisData.getActiveDatabase()); 
    g_controlExecutionTask2controlHashMap = restoreHashMapFromProperty("g_controlExecutionTask2controlHashMap", ArisData.getActiveDatabase());
    g_controlExecutionTask2controlExecutionOwnerHashMap = restoreHashMapFromProperty("g_controlExecutionTask2controlExecutionOwnerHashMap", ArisData.getActiveDatabase());
    g_controlExecutionTask2orgunitHashMap = restoreHashMapFromProperty("g_controlExecutionTask2orgunitHashMap", ArisData.getActiveDatabase());
    g_testdefinition2orgunitHashMap = restoreHashMapFromProperty("g_testdefinition2orgunitHashMap", ArisData.getActiveDatabase());
    g_testdefinition2controlsHashMap = restoreHashMapFromProperty("g_testdefinition2controlsHashMap", ArisData.getActiveDatabase());   
    g_testdefinition2testManagerHashMap = restoreHashMapFromProperty("g_testdefinition2testManagerHashMap", ArisData.getActiveDatabase());
	g_testdefinition2testerHashMap = restoreHashMapFromProperty("g_testdefinition2testerHashMap", ArisData.getActiveDatabase());
    g_testdefinition2testReviewerHashMap = restoreHashMapFromProperty("g_testdefinition2testReviewerHashMap", ArisData.getActiveDatabase());

	g_hierarchy2SignoffOwnerHashMap = restoreHashMapFromProperty("g_hierarchy2SignoffOwnerHashMap", ArisData.getActiveDatabase());
    g_hierarchy2TesterGroupHashMap = restoreHashMapFromProperty("g_hierarchy2TesterGroupHashMap", ArisData.getActiveDatabase());
    g_testerGroup2TesterHierarchyHashMap = restoreHashMapFromProperty("g_testerGroup2TesterHierarchyHashMap", ArisData.getActiveDatabase());
  
	g_hierarchy2TestAuditorHashMap = restoreHashMapFromProperty("g_hierarchy2TestAuditorHashMap", ArisData.getActiveDatabase());
	g_hierarchy2HierarchyOwnerHashMap = restoreHashMapFromProperty("g_hierarchy2HierarchyOwnerHashMap", ArisData.getActiveDatabase());
    
	return true;
}

/**
  * Store all test management HashSets and HashMaps
  */
function storeTestManagementToReportContext() {
    
    if ( !g_testManagementHashSetsFilled ) {return;}
    
    Context.setProperty("g_testManagementHashSetsFilled", "true");
    
    storeStringHashSetAsProperty("g_deactivatedLinkSet", g_deactivatedLinkSet);
    storeHashSetAsProperty("g_userGroupSet", g_userGroupSet);
    storeHashSetAsProperty("g_userSet", g_userSet);
    storeHashSetAsProperty("g_riskSet", g_riskSet);
    storeHashSetAsProperty("g_controlSet", g_controlSet);
    storeHashSetAsProperty("g_controlExecutionTaskSet", g_controlExecutionTaskSet);
    storeHashSetAsProperty("g_testdefinitionSet", g_testdefinitionSet);    
    storeHashSetAsProperty("g_processedRisksTestMode", g_processedRisksTestMode);
	storeHashSetAsProperty("g_processedControlsTestMode", g_processedControlsTestMode);
    storeHashSetAsProperty("g_error_risk_genericSpecific", g_error_risk_genericSpecific);
    storeHashSetAsProperty("g_testerSet", g_testerSet);
    storeHashSetAsProperty("g_regulationSet", g_regulationSet);
    storeHashSetAsProperty("g_orgUnitSet", g_orgUnitSet);
    storeHashSetAsProperty("g_processSet", g_processSet);
    
    storeHashMapAsProperty("g_risk2controlsHashMap", g_risk2controlsHashMap);
    storeHashMapAsProperty("g_risk2processesHashMap", g_risk2processesHashMap);
    storeHashMapAsProperty("g_risk2riskManagerHashMap", g_risk2riskManagerHashMap);
    storeHashMapAsProperty("g_risk2regulationsHashMap", g_risk2regulationsHashMap);
    storeHashMapAsProperty("g_risk2orgunitHashMap", g_risk2orgunitHashMap);
    storeHashMapAsProperty("g_specificRisk2genericRiskHashMap", g_specificRisk2genericRiskHashMap);
	storeHashMapAsProperty("g_control2processesHashMap", g_control2processesHashMap);
	storeHashMapAsProperty("g_control2regulationsHashMap", g_control2regulationsHashMap);
    storeHashMapAsProperty("g_control2risksHashMap", g_control2risksHashMap);
    storeHashMapAsProperty("g_control2testdefinitionsHashMap", g_control2testdefinitionsHashMap);
    storeHashMapAsProperty("g_control2controlManagerHashMap", g_control2controlManagerHashMap);
    storeHashMapAsProperty("g_control2controlExecutionTasksHashMap", g_control2controlExecutionTasksHashMap);
    storeHashMapAsProperty("g_controlExecutionTask2controlHashMap", g_controlExecutionTask2controlHashMap);
    storeHashMapAsProperty("g_controlExecutionTask2controlExecutionOwnerHashMap", g_controlExecutionTask2controlExecutionOwnerHashMap);
    storeHashMapAsProperty("g_controlExecutionTask2orgunitHashMap", g_controlExecutionTask2orgunitHashMap);   
    storeHashMapAsProperty("g_testdefinition2orgunitHashMap", g_testdefinition2orgunitHashMap);
    storeHashMapAsProperty("g_testdefinition2controlsHashMap", g_testdefinition2controlsHashMap);   
    storeHashMapAsProperty("g_testdefinition2testManagerHashMap", g_testdefinition2testManagerHashMap);
	storeHashMapAsProperty("g_testdefinition2testerHashMap", g_testdefinition2testerHashMap);
    storeHashMapAsProperty("g_testdefinition2testReviewerHashMap", g_testdefinition2testReviewerHashMap);

	storeHashMapAsProperty("g_hierarchy2SignoffOwnerHashMap", g_hierarchy2SignoffOwnerHashMap);
    storeHashMapAsProperty("g_hierarchy2TesterGroupHashMap", g_hierarchy2TesterGroupHashMap);
    storeHashMapAsProperty("g_testerGroup2TesterHierarchyHashMap", g_testerGroup2TesterHierarchyHashMap);
	
	storeHashMapAsProperty("g_hierarchy2TestAuditorHashMap", g_hierarchy2TestAuditorHashMap);
    storeHashMapAsProperty("g_hierarchy2HierarchyOwnerHashMap", g_hierarchy2HierarchyOwnerHashMap);
}
    
/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    ~~~~~~~ Determine start objects ~~~~~~~
                        ~~~~~~~ Test management ~~~~~~~
  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
/*---------------------------------------------------------------------------------------
    Determine the start objects relevant for basic ARCM test management depending on 
    the selected context and the approach:

	---- RBA ----
    - Group: 
    - Model: Risk diagram, Technical term model, Organisational chart 
    - ObjDef: Risk (RBA), Technical Term (Regulation hierarchy), OrgUnit (OrgUnit and
			  Tester hierarchy), Person type (User group)
	Processes the found risks only if the global flag "g_bTestManagementExport" is true
	(see determineMode() in arcm-common.js).
	
	---- CBA ----
	- Group: 
    - Model: Technical term model, Organisational chart 
    - ObjDef: (Function) (Control - CBA), Technical Term (Regulation hierarchy), OrgUnit 
			  (OrgUnit and Tester hierarchy), Person type (User group)
	Processes the found controls only if the global flag "g_bTestManagementExport" is true
	(see determineMode() in arcm-common.js).
	
    Fills the corresponding global HashSets.
    Returns an array of all found relevant objects.
 ---------------------------------------------------------------------------------------*/   
function getTestManagementSelectionByContext() {
    
    Context.writeStatus("Searching test management data");
    //start object search only if the object GUIDs are not already passed as Context properties
    if (!restoreTestManagementFromReportContext()) {
        buildAllowedModels();        
        // context "Group"
        var oSelectedGroups = ArisData.getSelectedGroups();
        if(oSelectedGroups.length == 0){
            var databases = ArisData.getSelectedDatabases();
            if (databases.length != 0) {
                oSelectedGroups = [databases[0].RootGroup()];
            }
        }
        if (oSelectedGroups.length > 0) {
            for (var i = 0; i < oSelectedGroups.length; i++ ) {
                getTestManagementRelevantObjects(oSelectedGroups[i])
            }
        }
        
        // context "Model"    
        var oSelectedModels = ArisData.getSelectedModels();
        if (oSelectedModels.length > 0) {
            for (var i = 0; i < oSelectedModels.length; i++ ) {
                
                //model must be marked as export relevant
                if (isExportRelevant(oSelectedModels[i])) {                    
                    classifyObjectsByTestManagementModel(oSelectedModels[i]);
                }
            }
        }
        
        // context "ObjDef"    
        var oSelectedObjDefs = ArisData.getSelectedObjDefs();
        if (oSelectedObjDefs.length > 0) {
            for (var i = 0; i < oSelectedObjDefs.length; i++ ) {
                if (g_approach == "RBA") {
                    if (oSelectedObjDefs[i].TypeNum() == Constants.OT_RISK) {
                        processTestManagementRisk(oSelectedObjDefs[i]);
                    }
                }
                if (g_approach == "CBA") {
                    if (isControl(oSelectedObjDefs[i])) {
                        processTestManagementControl(oSelectedObjDefs[i]);
                    }
                }
            }
        }
        
        //after global HashSets are filled:
        //Find all users to all found user groups   
        classifyUsers();
        removeEntriesFromUrgUnitSetIfinTesterSet();        
        g_testManagementHashSetsFilled = true;
        storeTestManagementToReportContext();
    }
}


/*---------------------------------------------------------------------------------------
    ---- RBA ----
	Determine the relevant objects in the given group (RBA -> risks)
	---- CBA ----
	Determine the relevant objects in the given group (CBA -> controls)
---------------------------------------------------------------------------------------*/
function getTestManagementRelevantObjects(p_oGroup) {
    
	if (g_approach == "RBA") {
		// Get export relevant risk objects on group
		var oRisks = p_oGroup.ObjDefListFilter(Constants.OT_RISK);
		for (var i = 0; i < oRisks.length; i++ ) {                
			processTestManagementRisk(oRisks[i]);
		}
	}
	
	if (g_approach == "CBA") {
		// Get export relevant control objects on group
		var oControls = p_oGroup.ObjDefListFilter(Constants.OT_FUNC);
		for (var i = 0; i < oControls.length; i++ ) {                
			if (isControl(oControls[i])) {
				processTestManagementControl(oControls[i]);
			}
		}
	}
    
    // Get objects in export relevant models
    var oModels = p_oGroup.ModelList();
    for (var i = 0; i < oModels.length; i++ ) {

        //model must be marked as export relevant        
        if (isExportRelevant(oModels[i])) {
            //aRelevantObjects = aRelevantObjects.concat(oModels[i].ObjDefList());
            classifyObjectsByTestManagementModel(oModels[i]);
        }
    }
    
    //Search by recursion in child groups
    var oChildGroups = p_oGroup.Childs();
    for (var i = 0; i < oChildGroups.length; i++ ) {
        getTestManagementRelevantObjects(oChildGroups[i]);
    }
    
}

/*---------------------------------------------------------------------------------------
    ---- CBA ----
	Determines for a export-relevant control which risks it has assigned; these risks
    are checked if they are generic or specific and processed accordingly
    (in test management mode).
    Then all other objects assigned to the control are handled by getFurtherTestManagementObjectsByControl(...)
 ---------------------------------------------------------------------------------------*/  
function processTestManagementControl(p_oControlDef) {

	//if the global flag "g_bTestManagementExport" is not set then don't process anything
	if (!g_bTestManagementExport) {
		return;
	}
	
    //if control was handled before in test management mode, skip it, else process it further
    if ( !g_processedControlsTestMode.add(p_oControlDef) ) {
        return;    
    }
    
    //control must be marked as export relevant
    if (isExportRelevant(p_oControlDef)) {
	
        //update global HashMap for Controls 
	    var aRiskDefs = getConnectedObjects(p_oControlDef, Constants.EDGES_IN, Constants.CT_IS_REDU_BY, Constants.OT_RISK);  
	    g_control2risksHashMap.put(p_oControlDef, aRiskDefs);
	    
	    //process all linked risks separately because the can be generic / specific
	    for (var i=0; i<aRiskDefs.length; i++ ) {
		
			//update global HashMap for linked risk 
			var linkedControls = g_risk2controlsHashMap.get(aRiskDefs[i]);
			if (linkedControls == null) {linkedControls = new Array();}
			linkedControls.push(p_oControlDef);
			g_risk2controlsHashMap.put(aRiskDefs[i], linkedControls);
			
			processTestManagementRisk(aRiskDefs[i]);        
	    }
	    
	    //since control is processed, add it to the global control set
	    g_controlSet.add(p_oControlDef);
	    
	    //process further objects linked to control
	    getFurtherTestManagementObjectsByControl(p_oControlDef);
    
    }
}

/*---------------------------------------------------------------------------------------   
	---- RBA ----
	Determines for a export-relevant risk if it is generic or specific or both 
    (in test management mode).
    1. If both: the risk is noted as erroneus
    2. If generic: only its assigned specific risks are processed further in test management mode
    3. If specific: the risk itself is processed further in test management mode
 ---------------------------------------------------------------------------------------*/  
function processTestManagementRisk(p_oRiskDef) {

	//if the global flag "g_bTestManagementExport" is not set then don't process anything
	if (!g_bTestManagementExport) {
		return;
	}

    //if risk was handled before in test management mode, skip it, else process it further
    if ( !g_processedRisksTestMode.add(p_oRiskDef) ) {
        return;    
    }
    
    //RBA - risk must be marked as export relevant
    //CBA - risk export relevant flag does not matter because this function was called by processTestManagementControl() 
    if ( (g_approach == "RBA" && isExportRelevant(p_oRiskDef))
            || g_approach == "CBA" ) {
            
        var aSpecializedRiskDefs = getConnectedObjects(p_oRiskDef, Constants.EDGES_OUT, Constants.CT_IS_ASSIG_6, Constants.OT_RISK);
        var aGenericRiskDefs = getConnectedObjects(p_oRiskDef, Constants.EDGES_IN, Constants.CT_IS_ASSIG_6, Constants.OT_RISK);             
            
        //1. A risk may only be generic (aSpecializedRiskDefs > 0) or specific (aGenericRiskDefs > 0)
        //If it is both:
        if (aSpecializedRiskDefs.length > 0 && aGenericRiskDefs.length > 0) {
            //note the risk as erroneous
            g_error_risk_genericSpecific.add(p_oRiskDef);
            //ignore this risk for further object collection
            return;
        }
            
        //2. if risk is itself generic
        if (aSpecializedRiskDefs.length > 0) {
                
            //Find connected objects for the specific risks...
            for (var j=0; j<aSpecializedRiskDefs.length; j++) {             
                
                g_specificRisk2genericRiskHashMap.put(aSpecializedRiskDefs[j], p_oRiskDef);
		        
                // process the specific risk
                processTestManagementRisk(aSpecializedRiskDefs[j]);  
            }    
            //... skip the generic risk itself
            return;
        }
            
        //3. if risk is itself specific
        if (aGenericRiskDefs.length > 0) {
                
            //if assigned generic exists and is not marked as export relevant too -> do not add both risks to the global risk set
            if (g_approach == "RBA" && !isExportRelevant(aGenericRiskDefs[0])) {
                return;    
            }
                
            g_specificRisk2genericRiskHashMap.put(p_oRiskDef, aGenericRiskDefs[0]);   
        }
        
        //since risk is processed, add it to the global risk set
        g_riskSet.add(p_oRiskDef);      
        // Find connected objects
		// RBA: Func, TechTerm, PersType, Ctrl, TestDef, OrgUnit
		// CBA: PersType
        getFurtherTestManagementObjectsByRisk(p_oRiskDef);        
    }   
}

/*---------------------------------------------------------------------------------------
    ---- RBA ----
	Determines the objects that are linked to the given risk:
    - Function
    - Technical term
    - Person type
    - Control
	---- CBA ----
	Determines the objects that are linked to the given risk:
    - Person type
 ---------------------------------------------------------------------------------------*/   
function getFurtherTestManagementObjectsByRisk(p_oSpecificRiskDef) {
    
    //-- own objects
    // Risk --> Func
    var aFuncDefs = getConnectedObjects(p_oSpecificRiskDef, Constants.EDGES_OUT, Constants.CT_OCCUR, Constants.OT_FUNC);  
    // Risk --> TechTerm
    var aRegulationsDefs = new Array();
    var aTechTermDefs = getConnectedObjects(p_oSpecificRiskDef, Constants.EDGES_OUT, Constants.CT_AFFECTS, Constants.OT_TECH_TRM);
    for (var i=0; i<aTechTermDefs.length; i++) {
        if (isRegulationItem(aTechTermDefs[i])) {
            aRegulationsDefs.push(aTechTermDefs[i]);
        }
    }
    // Risk --> PersType
    var aPersTypeDefs = new Array();
    aPersTypeDefs = aPersTypeDefs.concat( getConnectedObjects(p_oSpecificRiskDef, Constants.EDGES_IN, Constants.CT_IS_TECH_RESP_3, Constants.OT_PERS_TYPE) );
    aPersTypeDefs = aPersTypeDefs.concat( getConnectedObjects(p_oSpecificRiskDef, Constants.EDGES_IN, Constants.CT_IS_UNDER_RESP_OF, Constants.OT_PERS_TYPE) );
    var aAssignedRiskManagers = new Array();
    for (var j=0; j<aPersTypeDefs.length; j++) {
        if (isPersonTypeRiskManager(aPersTypeDefs[j]))  {aAssignedRiskManagers.push(aPersTypeDefs[j]);}
    } 
    // Risk --> Ctrl
	var aCtrlDefs = new Array();
	if (g_approach == "RBA") {
		aCtrlDefs = getConnectedObjects(p_oSpecificRiskDef, Constants.EDGES_OUT, Constants.CT_IS_REDU_BY, Constants.OT_FUNC);
	}
    
    
    //-- inherit objects
    //check if a generic risk exists to this specific risk
    var oGenericRisk = g_specificRisk2genericRiskHashMap.get(p_oSpecificRiskDef);
    
    //if so get the following linked risk management objects at the generic risk if there exist none for the specific risk:
    //  - regulations
    //  - risk manager
    if (oGenericRisk != null) {
    
	//Risk -> Func
        if (aFuncDefs.length == 0) {
            aFuncDefs = getConnectedObjects(oGenericRisk, Constants.EDGES_OUT, Constants.CT_OCCUR, Constants.OT_FUNC);    
        }
	
        // Risk -> TechTerm
        if (aRegulationsDefs.length == 0) {
            var aTechTermDefs = getConnectedObjects(oGenericRisk, Constants.EDGES_OUT, Constants.CT_AFFECTS, Constants.OT_TECH_TRM);
            for (var i=0; i<aTechTermDefs.length; i++) {
                if (isRegulationItem(aTechTermDefs[i])) {
                    aRegulationsDefs.push(aTechTermDefs[i]);
                }
            }
        }
        
        // Risk --> PersType
        if (aAssignedRiskManagers.length == 0) {
            aPersTypeDefs = new Array();
            aPersTypeDefs = aPersTypeDefs.concat( getConnectedObjects(oGenericRisk, Constants.EDGES_IN, Constants.CT_IS_TECH_RESP_3, Constants.OT_PERS_TYPE) );
            aPersTypeDefs = aPersTypeDefs.concat( getConnectedObjects(oGenericRisk, Constants.EDGES_IN, Constants.CT_IS_UNDER_RESP_OF, Constants.OT_PERS_TYPE) );
            //differentiate the roles
            for (var j=0; j<aPersTypeDefs.length; j++) {
                if (isPersonTypeRiskManager(aPersTypeDefs[j]))  {aAssignedRiskManagers.push(aPersTypeDefs[j]);}
            }
        }    
    }
    
    if (g_approach == "RBA") {
		//update global HashMap for Controls
		g_risk2controlsHashMap.put(p_oSpecificRiskDef, aCtrlDefs);
	}
    
    //update global HashMap for Processes
    g_risk2processesHashMap.put(p_oSpecificRiskDef, aFuncDefs);
     
    for (var i = 0; i < aFuncDefs.length; i++ ) {
        // Func --> ...
        var oFuncDef = aFuncDefs[i];
        // add current process to global set
        // if process was not handled before
        getFurtherObjectsByFunc(oFuncDef);
    }
    
    //update global HashMap for Regulations
    g_risk2regulationsHashMap.put(p_oSpecificRiskDef, aRegulationsDefs);  
    for (var i=0; i<aRegulationsDefs.length ;i++) {
        // add current regulations to global set
        // if regulation was not handled before
        getFurtherObjectsByRegulation(aRegulationsDefs[i]);
    } 
    
    //update global HashMap for Risk Managers
    g_risk2riskManagerHashMap.put(p_oSpecificRiskDef, aAssignedRiskManagers); 
    for (var i=0; i<aAssignedRiskManagers.length ;i++) {g_userGroupSet.add(aAssignedRiskManagers[i]);} // add group to global set
    
	if (g_approach == "RBA") {	
		//handle assigned Controls
		for (var i=0; i<aCtrlDefs.length; i++) {
			// Ctrl --> ...
			var oCtrlDef = aCtrlDefs[i];
			var linkedRisks = g_control2risksHashMap.get(oCtrlDef);
			if (linkedRisks == null) {linkedRisks = new Array();}
			linkedRisks.push(p_oSpecificRiskDef);
			g_control2risksHashMap.put(oCtrlDef, linkedRisks);
			// add current control to global set
			// if control was not handled before
			if ( g_controlSet.add(oCtrlDef) ) {                    
				//update global HashMap
				getFurtherTestManagementObjectsByControl(oCtrlDef);
			}
		}
	}
		
}

/*---------------------------------------------------------------------------------------
    ---- RBA ----
	Determines the objects that are linked to the given control:
    - Control execution tasks
    - Testdefinition
    - Person type
	---- CBA ----
	Determines the objects that are linked to the given control: 
    - Control execution tasks
    - Testdefinition
    - Person type   
    - Function
    - Technical term
 ---------------------------------------------------------------------------------------*/   
 function getFurtherTestManagementObjectsByControl(p_oCtrlDef) {

    // Ctrl --> Ctrl execution task
    var aCtrlExecutionTasks = getConnectedObjects(p_oCtrlDef, Constants.EDGES_OUT, Constants.CT_IS_INITIATED_BY, Constants.OT_CRTL_EXECUTION_TASK);
    //update global HashMap
    g_control2controlExecutionTasksHashMap.put(p_oCtrlDef, aCtrlExecutionTasks);      
    
   for (var i = 0; i < aCtrlExecutionTasks.length; i++ ) {              
        // First save linked controls of a control execution task
        var linkedControls = g_controlExecutionTask2controlHashMap.get(aCtrlExecutionTasks[i]);
        if (linkedControls == null) {linkedControls = new Array();}
        linkedControls.push(p_oCtrlDef);
        g_controlExecutionTask2controlHashMap.put(aCtrlExecutionTasks[i], linkedControls);
            
        // Ctrl execution task --> ...                          
        // add current control excecution task to global set
        // if control execution task was not handled before
        if ( g_controlExecutionTaskSet.add(aCtrlExecutionTasks[i]) ){            
            getFurtherObjectsByCtrlExecutionTask(aCtrlExecutionTasks[i]);
        }
    }    
    
    // Ctrl --> Testdef
    var aTestDefs = getConnectedObjects(p_oCtrlDef, Constants.EDGES_OUT, Constants.CT_IS_MONITORED_BY, Constants.OT_TEST_DEFINITION);
    //update global HashMap
    g_control2testdefinitionsHashMap.put(p_oCtrlDef, aTestDefs);  
    
	// Handle assigned testdefinitions
	for (var i = 0; i < aTestDefs.length; i++ ) {
        // Test --> ...                         
        //update global HashMap
        var linkedControls = g_testdefinition2controlsHashMap.get(aTestDefs[i]);
        if (linkedControls == null) {linkedControls = new Array();}
        linkedControls.push(p_oCtrlDef);
        g_testdefinition2controlsHashMap.put(aTestDefs[i], linkedControls);        
        // add current testdefinition to global set
        // if testdefinition was not handled before
        if ( g_testdefinitionSet.add(aTestDefs[i]) ){    
            getFurtherObjectsByTest(aTestDefs[i]);
        }
    }
    
    // Ctrl --> PersType
    var aPersTypeDefs = getConnectedObjects(p_oCtrlDef, Constants.EDGES_IN, Constants.CT_IS_TECH_RESP_3, Constants.OT_PERS_TYPE);
    var aAssignedControlManagers = new Array();
    for (var h=0; h<aPersTypeDefs.length; h++) {
        if (isPersonTypeControlManager(aPersTypeDefs[h])) {
            aAssignedControlManagers.push(aPersTypeDefs[h]);
        }
    }
    //update global HashMap
    g_control2controlManagerHashMap.put(p_oCtrlDef, aAssignedControlManagers);
    for (var i=0; i<aAssignedControlManagers.length ;i++) {g_userGroupSet.add(aAssignedControlManagers[i]);} // add groups to global set

	if (g_approach == "CBA") {
		// Control --> Func
		var aFuncDefs = getConnectedObjects(p_oCtrlDef, Constants.EDGES_OUT, Constants.CT_IS_PERFORMED_AT, Constants.OT_FUNC);  
		//update global HashMap for Processes
		g_control2processesHashMap.put(p_oCtrlDef, aFuncDefs);
		
		// Handle assigned functions as processes
		for (var i = 0; i < aFuncDefs.length; i++ ) {
			// Func --> ...
			var oFuncDef = aFuncDefs[i];
			// add current process to global set
			// if process was not handled before
            getFurtherObjectsByFunc(oFuncDef);
		}
		
		
		// Control --> TechTerm
		var aRegulationDefs = new Array();
		var aTechTermDefs = getConnectedObjects(p_oCtrlDef, Constants.EDGES_OUT, Constants.CT_AFFECTS, Constants.OT_TECH_TRM);
		for (var i=0; i<aTechTermDefs.length; i++) {
			if (isRegulationItem(aTechTermDefs[i])) {
				aRegulationDefs.push(aTechTermDefs[i]);
			}
		}
		//update global HashMap for Regulations
		g_control2regulationsHashMap.put(p_oCtrlDef, aTechTermDefs);  
		for (var i=0; i<aTechTermDefs.length ;i++) {
            // add current regulations to global set
            getFurtherObjectsByRegulation(aTechTermDefs[i]);
        } 
    }
}


/*---------------------------------------------------------------------------------------
    Determines the objects that are linked to the given control execution task;
    - Control owner group
    - Effected Orgunit
 ---------------------------------------------------------------------------------------*/   
function getFurtherObjectsByCtrlExecutionTask(p_oCtrlExecutionTaskDef) {

    // Ctrl --> Ctrl owner group
    var aCtrlExecutionTaskOwner = getConnectedObjects(p_oCtrlExecutionTaskDef, Constants.EDGES_IN, Constants.CT_IS_ASSIG_1, Constants.OT_PERS_TYPE);    
    var aAssignedControlExecutionOwners = new Array();
    for (var i=0; i<aCtrlExecutionTaskOwner.length; i++) {
        if (isPersonTypeControlExecutionOwner(aCtrlExecutionTaskOwner[i])) {
            aAssignedControlExecutionOwners.push(aCtrlExecutionTaskOwner[i]);
        }
    }    
    //update global HashMap
    g_controlExecutionTask2controlExecutionOwnerHashMap.put(p_oCtrlExecutionTaskDef, aAssignedControlExecutionOwners);      
    //update global set
    for (var i=0; i<aAssignedControlExecutionOwners.length ;i++) {g_userGroupSet.add(aAssignedControlExecutionOwners[i]);}
    
    // Ctrl --> Effected orgunit
    var aCtrlExecutionTaskOrgUnit = getConnectedObjects(p_oCtrlExecutionTaskDef, Constants.EDGES_OUT, Constants.CT_AFFECTS, Constants.OT_ORG_UNIT);
    
    //update global HashMap
    g_controlExecutionTask2orgunitHashMap.put(p_oCtrlExecutionTaskDef, aCtrlExecutionTaskOrgUnit);   
    //update global set
    for (var i=0; i<aCtrlExecutionTaskOrgUnit.length ;i++) {
        // add organisational units to global set
        getFurtherOrgUnitHierarchyObjectsByOrgUnitDef(aCtrlExecutionTaskOrgUnit[i]);
    }
}


/*---------------------------------------------------------------------------------------
    Determines the objects that are linked to the given testdefinition:
    - Organisational unit
    - Person type
 ---------------------------------------------------------------------------------------*/   
function getFurtherObjectsByTest(p_oTestDef) {

    // Test --> OrgUnit
    var aOrgUnitDefs = getConnectedObjects(p_oTestDef, Constants.EDGES_OUT, Constants.CT_AFFECTS, Constants.OT_ORG_UNIT);
    //update global HashMap
    g_testdefinition2orgunitHashMap.put(p_oTestDef, aOrgUnitDefs);
    for (var i=0; i<aOrgUnitDefs.length ;i++) {
        // add organisational units to global set
        getFurtherOrgUnitHierarchyObjectsByOrgUnitDef(aOrgUnitDefs[i]);
    } 

    // Test --> PersType
    var aPersTypeDefs = getConnectedObjects(p_oTestDef, Constants.EDGES_IN, Constants.CT_IS_ASSIG_1, Constants.OT_PERS_TYPE);
    var aAssignedTestManagers = new Array();
	var aAssignedTesters = new Array();
    var aAssignedTestReviewers = new Array();
    for (var i=0; i<aPersTypeDefs.length ;i++) {
		if (isPersonTypeTestManager(aPersTypeDefs[i])) {
            aAssignedTestManagers.push(aPersTypeDefs[i]);
        }
        if (isPersonTypeTester(aPersTypeDefs[i])) {
            aAssignedTesters.push(aPersTypeDefs[i]);
        }
        if (isPersonTypeTestReviewer(aPersTypeDefs[i])) {
            aAssignedTestReviewers.push(aPersTypeDefs[i]);
        }
    }
    
	//update global HashMap
    g_testdefinition2testManagerHashMap.put(p_oTestDef, aAssignedTestManagers);
    for (var m=0; m<aAssignedTestManagers.length; m++) {
        g_userGroupSet.add(aAssignedTestManagers[m]);   
    }
	
    //update global HashMap
    g_testdefinition2testerHashMap.put(p_oTestDef, aAssignedTesters);
    for (var k=0; k<aAssignedTesters.length; k++) {
        g_userGroupSet.add(aAssignedTesters[k]);
        //lookup all connected tester hierarchy objects
        var aTesterHierarchyDefs = getConnectedObjects(aAssignedTesters[k], Constants.EDGES_OUT, Constants.CT_WRK_IN, Constants.OT_ORG_UNIT);
        for (var x=0; x<aTesterHierarchyDefs.length; x++) {
            getFurtherTesterHierarchyObjectsByOrgUnitDef(aTesterHierarchyDefs[x]);
        }
    }
    //update global HashMap
    g_testdefinition2testReviewerHashMap.put(p_oTestDef, aAssignedTestReviewers);
    for (var m=0; m<aAssignedTestReviewers.length; m++) {
        g_userGroupSet.add(aAssignedTestReviewers[m]);   
    }     
}


/*---------------------------------------------------------------------------------------
    Checks for objects of type "Organisational Unit" if they are part of the orgunit
    hierarchy for the test management mode. It is if it has one or more assigned 
    testdefinitions.
 ---------------------------------------------------------------------------------------*/
function isOrgunitInTestManagementOrgUnitHierarchy(p_orgUnitDef) {
    
    var aTestdefinitionDefs = getConnectedObjects(p_orgUnitDef, Constants.EDGES_IN, Constants.CT_AFFECTS, Constants.OT_TEST_DEFINITION);     
    return aTestdefinitionDefs.length > 0;    
}    
    
/*---------------------------------------------------------------------------------------
    Checks for objects of type "Organisational Unit" if they are part of the tester
    hierarchy. It is if it has exactly one assigned role object with level "Tester".
 ---------------------------------------------------------------------------------------*/        
function isOrgunitInTesterHierarchy(p_userGroupDef) {
  
    var bPartOfTesterHierarchy = false;  
    //determine connected usergroups
    var aUsergroupDefs = getConnectedObjects(p_userGroupDef, Constants.EDGES_IN, Constants.CT_WRK_IN, Constants.OT_PERS_TYPE);
    for (var i=0; i<aUsergroupDefs.length && !bPartOfTesterHierarchy; i++) {
        bPartOfTesterHierarchy = isPersonTypeInTestHierarchy(aUsergroupDefs[i]);
    }     
    return bPartOfTesterHierarchy;
    
}

/*---------------------------------------------------------------------------------------
    Checks for objects of type "User group" if they are part of the tester
    hierarchy. It is if 
    1. it has exactly one assigned role object which name starts with "Tester".
    2. it is also connected to a organisational unit (which will be the tester hierarchy element)
 ---------------------------------------------------------------------------------------*/        
function isPersonTypeInTestHierarchy(p_userGroupDef) {
  
    var bPartOfTesterHierarchy  = false;
    //if it is a tester group
    if ( isPersonTypeTester(p_userGroupDef) ) {
        var connectedRoles = getConnectedObjects(p_userGroupDef, Constants.EDGES_OUT, Constants.CT_WRK_IN, Constants.OT_ORG_UNIT);               
        if (connectedRoles.length > 0) {
            bPartOfTesterHierarchy = true;
        }
    }
    return bPartOfTesterHierarchy;
}

/*---------------------------------------------------------------------------------------
    Checks for objects of type "User group" if they are test managers.
    For standard mapping: They are if they have one assigned role object which name starts 
    with "Test manager".
 ---------------------------------------------------------------------------------------*/
function isPersonTypeTestManager(p_userGroupDef) {
    return hasARCMRole(p_userGroupDef, "testmanager");
}

/*---------------------------------------------------------------------------------------
    Checks for objects of type "User group" if they are testers.
    For standard mapping: They are if they have one assigned role object which name starts 
    with "Tester".
 ---------------------------------------------------------------------------------------*/
function isPersonTypeTester(p_userGroupDef) {
    return hasARCMRole(p_userGroupDef, "tester");
}
    
/*---------------------------------------------------------------------------------------
    Checks for objects of type "User group" if they are test reviewers.
    For standard mapping: They are if they have one assigned role object which name starts 
    with "Test reviewer".
 ---------------------------------------------------------------------------------------*/
function isPersonTypeTestReviewer(p_userGroupDef) {
    return hasARCMRole(p_userGroupDef, "testreviewer");
}

/*---------------------------------------------------------------------------------------
    Checks for objects of type "User group" if they are risk managers.
    For standard mapping: They are if they have one assigned role object which name starts 
    with "Risk manager".
 ---------------------------------------------------------------------------------------*/        
function isPersonTypeRiskManager(p_userGroupDef) {
    return hasARCMRole(p_userGroupDef, "riskmanager");
}

/*---------------------------------------------------------------------------------------
    Checks for objects of type "User group" if they are control managers.
    For standard mapping: They are if they have one assigned role object which name starts 
    with "Control manager".
 ---------------------------------------------------------------------------------------*/        
function isPersonTypeControlManager(p_userGroupDef) {
    return hasARCMRole(p_userGroupDef, "controlmanager");
}

/*---------------------------------------------------------------------------------------
    Checks for objects of type "User group" if they are control execution owners.
    For standard mapping: They are if they have one assigned role object which name starts 
    with "Control execution owner".
 ---------------------------------------------------------------------------------------*/        
function isPersonTypeControlExecutionOwner(p_userGroupDef) {
    return hasARCMRole(p_userGroupDef, "controlexecutionowner");
}

/*---------------------------------------------------------------------------------------
    Checks for objects of type "User group" if they are Test auditors.
    For standard mapping: They are if they have one assigned role object which name starts 
    with "Test auditor".
 ---------------------------------------------------------------------------------------*/    
function isPersonTypeTestAuditor(p_userGroupDef) {
    return hasARCMRole(p_userGroupDef, "testauditor");
}

/*---------------------------------------------------------------------------------------
    Checks for objects of type "User group" if they are Hierarchy Owners.
    For standard mapping: They are if they have one assigned role object which name starts 
    with "Hierarchy owner".
 ---------------------------------------------------------------------------------------*/    
function isPersonTypeHierarchyOwner(p_userGroupDef) {
    return hasARCMRole(p_userGroupDef, "hierarchyowner");
}

/*---------------------------------------------------------------------------------------
    Checks for objects of type "User group" if they are SignOff owners.
    For standard mapping: They are if they have one assigned role object which name starts 
    with "Sign-off owner".
 ---------------------------------------------------------------------------------------*/    
function isPersonTypeSignOffOwner(p_userGroupDef) {
    return hasARCMRole(p_userGroupDef, "signoffowner");
}


/*---------------------------------------------------------------------------------------
    Classify objects of a given model according to test management modeling
 ---------------------------------------------------------------------------------------*/
function classifyObjectsByTestManagementModel(p_oModel) {
    
    var aObjectsToClassify = p_oModel.ObjDefList();
    
    //risk diagrams
	if (g_approach == "RBA") {
		if (p_oModel.TypeNum() == Constants.MT_RISK_DGM) {
			for (var i=0; i<aObjectsToClassify.length; i++) {
				if (aObjectsToClassify[i].TypeNum() == Constants.OT_RISK) {
					processTestManagementRisk(aObjectsToClassify[i]);
				}
			}
		}
	}
        
    //organisational charts
    if (p_oModel.TypeNum() == Constants.MT_ORG_CHRT) {
        for (var i=0; i<aObjectsToClassify.length; i++) {
            //ABA organisational unit objects
            if (aObjectsToClassify[i].TypeNum() == Constants.OT_ORG_UNIT) {
                if ( isOrgunitInTesterHierarchy(aObjectsToClassify[i]) ) {
                    //connected tester group -> tester hierarchy
                    getFurtherTesterHierarchyObjectsByOrgUnitDef(aObjectsToClassify[i]);
                }
            }
            //person types -> user groups
            if (aObjectsToClassify[i].TypeNum() == Constants.OT_PERS_TYPE) {
                
                //only "real" usergroups, not their role objects
                if (getRoleObject(aObjectsToClassify[i]) == null) {
                    continue;
                }
                g_userGroupSet.add(aObjectsToClassify[i]);
                //if tester group -> add linked organisational unit to global tester hierarchy Set
                if (isPersonTypeInTestHierarchy(aObjectsToClassify[i])) {
                    var aTesterHierarchyDefs = getConnectedObjects(aObjectsToClassify[i], Constants.EDGES_OUT, Constants.CT_WRK_IN, Constants.OT_ORG_UNIT);
                    for (var j=0; j<aTesterHierarchyDefs.length ;j++) {
                        getFurtherTesterHierarchyObjectsByOrgUnitDef(aTesterHierarchyDefs[j]);
                    }    
                }
            }
        }    
    }
    getProcessesFromModel(p_oModel, aObjectsToClassify);
	getRegulationsFromModel(p_oModel, aObjectsToClassify);
    getOrgUnitsFromModel(p_oModel, aObjectsToClassify);
    getApplicationSystemTypesFromModel(p_oModel, aObjectsToClassify);    
}

/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ~~~~~~~ Restore HashSets and HashMaps from report context if possible ~~~~~~~
                        ~~~~~~~ Risk management ~~~~~~~
  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
  
/**
  * Restore all risk management HashSets and HashMaps
  */
function restoreRiskManagementFromReportContext() {
    
    if ( Context.getProperty("g_riskManagementHashSetsFilled") == null 
         || !Context.getProperty("g_riskManagementHashSetsFilled").equals("true") ) {
         return false;
    }
	
	//in case the following HashSets and HashMaps were not filled/restored during test management earlier
	if (!g_testManagementHashSetsFilled) {
		g_deactivatedLinkSet = restoreStringHashSetFromProperty("g_deactivatedLinkSet", ArisData.getActiveDatabase());
		g_userGroupSet = restoreHashSetFromProperty("g_userGroupSet", ArisData.getActiveDatabase());
		g_userSet = restoreHashSetFromProperty("g_userSet", ArisData.getActiveDatabase());
		g_riskSet = restoreHashSetFromProperty("g_riskSet", ArisData.getActiveDatabase());
		g_error_risk_genericSpecific = restoreHashSetFromProperty("g_error_risk_genericSpecific", ArisData.getActiveDatabase());
		g_regulationSet = restoreHashSetFromProperty("g_regulationSet", ArisData.getActiveDatabase());
		g_orgUnitSet = restoreHashSetFromProperty("g_orgUnitSet", ArisData.getActiveDatabase());
		g_processSet = restoreHashSetFromProperty("g_processSet", ArisData.getActiveDatabase());
		
		g_risk2processesHashMap = restoreHashMapFromProperty("g_risk2processesHashMap", ArisData.getActiveDatabase());
		g_risk2regulationsHashMap = restoreHashMapFromProperty("g_risk2regulationsHashMap", ArisData.getActiveDatabase());
		g_risk2orgunitHashMap = restoreHashMapFromProperty("g_risk2orgunitHashMap", ArisData.getActiveDatabase());
		g_risk2riskManagerHashMap = restoreHashMapFromProperty("g_risk2riskManagerHashMap", ArisData.getActiveDatabase());
	}
     
    g_processedRisksRiskMode = restoreHashSetFromProperty("g_processedRisksRiskMode", ArisData.getActiveDatabase());
    g_processedControlsRiskMode = restoreHashSetFromProperty("g_processedControlsRiskMode", ArisData.getActiveDatabase());
	g_appSysSet = restoreHashSetFromProperty("g_appSysSet", ArisData.getActiveDatabase());
    g_riskCategorySet = restoreHashSetFromProperty("g_riskCategorySet", ArisData.getActiveDatabase());
   
	g_risk2riskManagerHashMap = restoreHashMapFromProperty("g_risk2riskManagerHashMap", ArisData.getActiveDatabase());   
    g_risk2appSysTypeHashMap = restoreHashMapFromProperty("g_risk2appSysTypeHashMap", ArisData.getActiveDatabase());
    g_risk2riskCategoryHashMap = restoreHashMapFromProperty("g_risk2riskCategoryHashMap", ArisData.getActiveDatabase());
    g_risk2riskOwnerHashMap = restoreHashMapFromProperty("g_risk2riskOwnerHashMap", ArisData.getActiveDatabase());
    g_risk2riskReviewerHashMap = restoreHashMapFromProperty("g_risk2riskReviewerHashMap", ArisData.getActiveDatabase());
    
    return true;
}

/**
  * Store all risk management HashSets and HashMaps
  */
function storeRiskManagementToReportContext() {
    
    if (!g_riskManagementHashSetsFilled) {return;}

    Context.setProperty("g_riskManagementHashSetsFilled", "true");	
	
    //might be already stored during test management but we have to update them in case new objects were processed during risk management
    storeStringHashSetAsProperty("g_deactivatedLinkSet", g_deactivatedLinkSet);
    storeHashSetAsProperty("g_userGroupSet", g_userGroupSet);
    storeHashSetAsProperty("g_userSet", g_userSet);
    storeHashSetAsProperty("g_riskSet", g_riskSet);
    storeHashSetAsProperty("g_error_risk_genericSpecific", g_error_risk_genericSpecific);
    storeHashSetAsProperty("g_regulationSet", g_regulationSet);
    storeHashSetAsProperty("g_orgUnitSet", g_orgUnitSet);
    storeHashSetAsProperty("g_processSet", g_processSet);
    
	storeHashMapAsProperty("g_risk2processesHashMap", g_risk2processesHashMap);
	storeHashMapAsProperty("g_risk2regulationsHashMap", g_risk2regulationsHashMap);
	storeHashMapAsProperty("g_risk2orgunitHashMap", g_risk2orgunitHashMap);
	storeHashMapAsProperty("g_risk2riskManagerHashMap", g_risk2riskManagerHashMap);
	
    //the purely risk management relevant HashSets and HashMaps
    storeHashSetAsProperty("g_processedRisksRiskMode", g_processedRisksRiskMode);
	storeHashSetAsProperty("g_processedControlsRiskMode", g_processedControlsRiskMode);
    storeHashSetAsProperty("g_appSysSet", g_appSysSet);
    storeHashSetAsProperty("g_riskCategorySet", g_riskCategorySet);
    
	storeHashMapAsProperty("g_risk2riskManagerHashMap", g_risk2riskManagerHashMap);
    storeHashMapAsProperty("g_risk2appSysTypeHashMap", g_risk2appSysTypeHashMap);
    storeHashMapAsProperty("g_risk2riskCategoryHashMap", g_risk2riskCategoryHashMap);
    storeHashMapAsProperty("g_risk2riskOwnerHashMap", g_risk2riskOwnerHashMap);
    storeHashMapAsProperty("g_risk2riskReviewerHashMap", g_risk2riskReviewerHashMap);
}


/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    ~~~~~~~ Determine start objects ~~~~~~~
                        ~~~~~~~ Risk management ~~~~~~~
  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
/*---------------------------------------------------------------------------------------
    Determine the start objects relevant for basic ARCM risk management depending on 
    the selected context: 
    - Group:
	- Model: Risk diagram, Application system type diagram, Organisational chart,
			 Technical term model
    - ObjDef: Risk, Application system types, OrgUnit (OrgUnit and Tester hierarchy),
		      Person type (User group), Technical term (Regulation hierarchy)
	Processes the found risks only if the global flag "g_bRiskManagementExport" is true
	(see determineMode() in arcm-common.js).
    Fills the corresponding global HashSets.
    Returns an array of all found relevant objects.
 ---------------------------------------------------------------------------------------*/
function getRiskManagementSelectionByContext() {
	
    Context.writeStatus("Searching risk management data");
    //start object search only if the object GUIDs are not already passed as Context properties
    if (!restoreRiskManagementFromReportContext()) {
        buildAllowedModels();         
        // context "Group"
        var oSelectedGroups = ArisData.getSelectedGroups();
        if(oSelectedGroups.length == 0){
            var databases = ArisData.getSelectedDatabases();
            if (databases.length != 0) {
                oSelectedGroups = [databases[0].RootGroup()];
            }
        }
        if (oSelectedGroups.length > 0) {
            for (var i = 0; i < oSelectedGroups.length; i++ ) {
                getRiskManagementRelevantObjects(oSelectedGroups[i])
            }
        }
        
        // context "Model"    
        var oSelectedModels = ArisData.getSelectedModels();
        if (oSelectedModels.length > 0) {
            for (var i = 0; i < oSelectedModels.length; i++ ) {        
                //model must be marked as export relevant
                if (isExportRelevant(oSelectedModels[i])) {
                    classifyObjectsByRiskManagementModel(oSelectedModels[i])
                }
            }
        }
        
        // context "ObjDef"    
        var oSelectedObjDefs = ArisData.getSelectedObjDefs();
        if (oSelectedObjDefs.length > 0) {
            for (var i = 0; i < oSelectedObjDefs.length; i++ ) {
                if (g_approach == "RBA") {
                    if (oSelectedObjDefs[i].TypeNum() == Constants.OT_RISK) {
                        processRiskManagementRisk(oSelectedObjDefs[i]);
                    }
                }	
                if (g_approach == "CBA") {
                    if (isControl(oSelectedObjDefs[i])) {
                        processRiskManagementControl(oSelectedObjDefs[i]);
                    }
                    //CBA special - process also risks without export relevant controls if they are set to 'risk management' = 'true'
                    if ( oSelectedObjDefs[i].TypeNum() == Constants.OT_RISK
                         && isboolattributetrue(oSelectedObjDefs[i], Constants.AT_GRC_RISK_MANAGEMENT_RELEVANT, g_nLoc) ) {
                        processRiskManagementRisk(oSelectedObjDefs[i]);
                    }
                }	
            }
        }
        
        //after global HashSets are filled:   
        //Find all users to all found user groups
        classifyUsers();
        removeEntriesFromUrgUnitSetIfinTesterSet();
        
        g_riskManagementHashSetsFilled = true;
        storeRiskManagementToReportContext();
    }
}

/*---------------------------------------------------------------------------------------
    Determine the relevant objects in the given group:
	RBA -> risks
	CBA -> controls
---------------------------------------------------------------------------------------*/
function getRiskManagementRelevantObjects(p_oGroup) {
    
	if (g_approach == "RBA") {
		// Get export relevant risk objects
		var oRisks = p_oGroup.ObjDefListFilter(Constants.OT_RISK);
		for (var i = 0; i < oRisks.length; i++ ) {             
			processRiskManagementRisk(oRisks[i]);
		}
	}
	
	if (g_approach == "CBA") {
		// Get export relevant control objects on group
		var oControls = p_oGroup.ObjDefListFilter(Constants.OT_FUNC);
		for (var i = 0; i < oControls.length; i++ ) {                
			if (isControl(oControls[i])) {
				processRiskManagementControl(oControls[i]);
			}
		}
		//CBA special - process also risks without export relevant controls if they are set to 'risk management' = 'true'
		var oRisks = p_oGroup.ObjDefListFilter(Constants.OT_RISK);
		for (var i = 0; i < oRisks.length; i++ ) {
			if ( isboolattributetrue(oRisks[i], Constants.AT_GRC_RISK_MANAGEMENT_RELEVANT, g_nLoc) ) {
				processRiskManagementRisk(oRisks[i]);    
			}
		}
	}
	
    
    // Get objects in export relevant models
    var oModels = p_oGroup.ModelList();
    for (var i = 0; i < oModels.length; i++ ) {

        //model must be marked as export relevant        
        if (isExportRelevant(oModels[i])) {
            //aRelevantObjects = aRelevantObjects.concat(oModels[i].ObjDefList());
            classifyObjectsByRiskManagementModel(oModels[i]);
        }
    }
    
    //Search by recursion in child groups
    var oChildGroups = p_oGroup.Childs();
    for (var i = 0; i < oChildGroups.length; i++ ) {
        getRiskManagementRelevantObjects(oChildGroups[i]);
    }
    
}

/*---------------------------------------------------------------------------------------
	---- CBA ----
    Determines for a export-relevant control which risks it has assigned; these risks
    are checked if they are generic or specific and processed accordingly
    (in test management mode).
    Then all other objects assigned to the control are handled by getFurtherTestManagementObjectsByControl(...)
 ---------------------------------------------------------------------------------------*/  
function processRiskManagementControl(p_oControlDef) {

    //if control was handled before in risk management mode, skip it, else process it further
    if ( !g_processedControlsRiskMode.add(p_oControlDef) ) {
        return;    
    }
    
    //control must be marked as export relevant
    if (isExportRelevant(p_oControlDef)) {
    
        //update global HashMap for Controls 
        var aRiskDefs = getConnectedObjects(p_oControlDef, Constants.EDGES_IN, Constants.CT_IS_REDU_BY, Constants.OT_RISK);  
        g_control2risksHashMap.put(p_oControlDef, aRiskDefs);
        
        //process all linked risks separately because the can be generic / specific
        for (var i=0; i<aRiskDefs.length; i++ ) {
            
            //update global HashMap for linked risk 
            var linkedControls = g_risk2controlsHashMap.get(aRiskDefs[i]);
            if (linkedControls == null) {linkedControls = new Array();}
            linkedControls.push(p_oControlDef);
            g_risk2controlsHashMap.put(aRiskDefs[i], linkedControls);
            
            processRiskManagementRisk(aRiskDefs[i]);        
        }
        
        //since control is processed, add it to the global control set
        g_controlSet.add(p_oControlDef);
        
        //process further objects linked to control
        getFurtherRiskManagementObjectsByControl(p_oControlDef);
        
    }
}

/*---------------------------------------------------------------------------------------
    Determines for a export-relevant risk if it is generic or specific or both 
    (in risk management mode).
    1. If both: the risk is noted as erroneus
    2. If generic: only its assigned specific risks are processed further in risk management mode
    3. If specific: the risk itself is processed further in risk management mode
 ---------------------------------------------------------------------------------------*/  
function processRiskManagementRisk(p_oRiskDef) {

	//if the global flag "g_bRiskManagementExport" is not set then don't process anything
	if (!g_bRiskManagementExport) {
		return;
	}

    //if risk was handled before in risk management mode, skip it, else process it further
    if ( !g_processedRisksRiskMode.add(p_oRiskDef) ) {
        return;    
    }
    
    //risk must be marked as export relevant
    if (isExportRelevant(p_oRiskDef)) {
            
        var aSpecializedRiskDefs = getConnectedObjects(p_oRiskDef, Constants.EDGES_OUT, Constants.CT_IS_ASSIG_6, Constants.OT_RISK);
        var aGenericRiskDefs = getConnectedObjects(p_oRiskDef, Constants.EDGES_IN, Constants.CT_IS_ASSIG_6, Constants.OT_RISK);             
            
        //1. A risk may only be generic (aSpecializedRiskDefs > 0) or specific (aGenericRiskDefs > 0)
        //If it is both:
        if (aSpecializedRiskDefs.length > 0 && aGenericRiskDefs.length > 0) {
            //note the risk as erroneous
            g_error_risk_genericSpecific.add(p_oRiskDef);
            //ignore this risk for further object collection
            return;
        }
            
        //2. if risk is itself generic
        if (aSpecializedRiskDefs.length > 0) {
                
            //Find connected objects for the specific risks...
            for (var j=0; j<aSpecializedRiskDefs.length; j++) {
                
                g_specificRisk2genericRiskHashMap.put(aSpecializedRiskDefs[j], p_oRiskDef);

                // process the specific risk
                processRiskManagementRisk(aSpecializedRiskDefs[j]);                 
            }    
            //... skip the generic risk itself
            return;
        }
            
        //3. if risk is itself specific
        if (aGenericRiskDefs.length > 0) {
                
            //if assigned generic exists and is not marked as export relevant too -> do not add both risks to the global risk set
            if (!isExportRelevant(aGenericRiskDefs[0])) {
                return;    
            }
                
            g_specificRisk2genericRiskHashMap.put(p_oRiskDef, aGenericRiskDefs[0]);   
        }
       
        //since risk is processed, add it to the global risk set
        g_riskSet.add(p_oRiskDef);
    
        // Find connected objects (Func, TechTerm, PersType, Ctrl, TestDef, OrgUnit)
        getFurtherRiskManagementObjectsByRisk(p_oRiskDef);
        
    }   
}

/*---------------------------------------------------------------------------------------
	Determines the objects that are linked to the given risk:
    - Function
    - Technical term
    - Organisational units
    - Application system types
    - Person type
    - Risk category
 ---------------------------------------------------------------------------------------*/   
function getFurtherRiskManagementObjectsByRisk(p_oSpecificRiskDef) {
    
    //get all linked objects for current risk that are related to risk management
    //Risk -> Func
    var aFuncDefs = getConnectedObjects(p_oSpecificRiskDef, Constants.EDGES_OUT, Constants.CT_OCCUR, Constants.OT_FUNC);
    // Risk -> TechTerm
    var aRegulationDefs = new Array();
    var aTechTermDefs = getConnectedObjects(p_oSpecificRiskDef, Constants.EDGES_OUT, Constants.CT_AFFECTS, Constants.OT_TECH_TRM);
    for (var i=0; i<aTechTermDefs.length; i++) {
        if (isRegulationItem(aTechTermDefs[i])) {
            aRegulationDefs.push(aTechTermDefs[i]);
        }
    }
    // Risk -> OrgUnit
    var aOrgUnitDefs = getConnectedObjects(p_oSpecificRiskDef, Constants.EDGES_OUT, Constants.CT_AFFECTS, Constants.OT_ORG_UNIT);
    // Risk -> Application System Type
    var aAppSysTypeDefs = getConnectedObjects(p_oSpecificRiskDef, Constants.EDGES_OUT, Constants.CT_AFFECTS, Constants.OT_APPL_SYS_TYPE);
    // Test --> PersType
    var aPersTypeDefs = new Array();
    aPersTypeDefs = aPersTypeDefs.concat( getConnectedObjects(p_oSpecificRiskDef, Constants.EDGES_IN, Constants.CT_IS_TECH_RESP_3, Constants.OT_PERS_TYPE) );
    aPersTypeDefs = aPersTypeDefs.concat( getConnectedObjects(p_oSpecificRiskDef, Constants.EDGES_IN, Constants.CT_IS_UNDER_RESP_OF, Constants.OT_PERS_TYPE) );
            
    var aAssignedRiskManagers   = new Array();
    var aAssignedRiskOwners     = new Array();
    var aAssignedRiskReviewers  = new Array();
    //differentiate the roles
    for (var j=0; j<aPersTypeDefs.length; j++) {
        if (isPersonTypeRiskManager(aPersTypeDefs[j]))  {aAssignedRiskManagers.push(aPersTypeDefs[j]);}
        if (isPersonTypeRiskOwner(aPersTypeDefs[j]))    {aAssignedRiskOwners.push(aPersTypeDefs[j]);}
        if (isPersonTypeRiskReviewer(aPersTypeDefs[j])) {aAssignedRiskReviewers.push(aPersTypeDefs[j]);}
    }     
    // Risk --> Risk category
    var aRiskCategoryDefs = getConnectedObjects(p_oSpecificRiskDef, Constants.EDGES_IN, Constants.CT_SUBS_1, Constants.OT_RISK_CATEGORY);   
    
    
    //check if a generic risk exists to this specific risk
    var oGenericRisk = g_specificRisk2genericRiskHashMap.get(p_oSpecificRiskDef);
    
    //if so get all linked risk management objects at the generic risk if there exist none for the specific risk
    if (oGenericRisk != null) {
        
        //Risk -> Func
        if (aFuncDefs.length == 0) {
            aFuncDefs = getConnectedObjects(oGenericRisk, Constants.EDGES_OUT, Constants.CT_OCCUR, Constants.OT_FUNC);    
        } 
        // Risk -> TechTerm
        if (aTechTermDefs.length == 0) {
            aTechTermDefs = getConnectedObjects(oGenericRisk, Constants.EDGES_OUT, Constants.CT_AFFECTS, Constants.OT_TECH_TRM);
		    for (var i=0; i<aTechTermDefs.length; i++) {
				if (isRegulationItem(aTechTermDefs[i])) {
				    aRegulationDefs.push(aTechTermDefs[i]);
				}
		    }
        }
        // Risk -> OrgUnit
        if (aOrgUnitDefs.length == 0) {
            aOrgUnitDefs = getConnectedObjects(oGenericRisk, Constants.EDGES_OUT, Constants.CT_AFFECTS, Constants.OT_ORG_UNIT);
        }
        // Risk -> Application System Type
        if (aAppSysTypeDefs.length == 0) {
            aAppSysTypeDefs = getConnectedObjects(oGenericRisk, Constants.EDGES_OUT, Constants.CT_AFFECTS, Constants.OT_APPL_SYS_TYPE);
        }
        // Risk --> PersType
        if (aAssignedRiskManagers.length == 0 || aAssignedRiskOwners.length == 0 || aAssignedRiskReviewers == 0) {          
            aPersTypeDefs = new Array();
            aPersTypeDefs = aPersTypeDefs.concat( getConnectedObjects(oGenericRisk, Constants.EDGES_IN, Constants.CT_IS_TECH_RESP_3, Constants.OT_PERS_TYPE) );
            aPersTypeDefs = aPersTypeDefs.concat( getConnectedObjects(oGenericRisk, Constants.EDGES_IN, Constants.CT_IS_UNDER_RESP_OF, Constants.OT_PERS_TYPE) );
            
            //differentiate the roles, inherit only from generic risk if specific risk has no such groups already
            if (aAssignedRiskManagers.length == 0) {
                for (var j=0; j<aPersTypeDefs.length; j++) {
                    if (isPersonTypeRiskManager(aPersTypeDefs[j]))  {aAssignedRiskManagers.push(aPersTypeDefs[j]);}     
                }
            }
            if (aAssignedRiskOwners.length == 0) {
                for (var j=0; j<aPersTypeDefs.length; j++) {
                    if (isPersonTypeRiskOwner(aPersTypeDefs[j]))    {aAssignedRiskOwners.push(aPersTypeDefs[j]);}     
                }
            }
            if (aAssignedRiskReviewers.length == 0) {
                for (var j=0; j<aPersTypeDefs.length; j++) {
                    if (isPersonTypeRiskReviewer(aPersTypeDefs[j])) {aAssignedRiskReviewers.push(aPersTypeDefs[j]);}     
                }
            }  
        }
        
        // Risk --> Risk category
        if (aRiskCategoryDefs.length == 0) {
            aRiskCategoryDefs = getConnectedObjects(oGenericRisk, Constants.EDGES_IN, Constants.CT_SUBS_1, Constants.OT_RISK_CATEGORY);   
        }
        
    }// endif genericRisk
    
    	
	//update global HashMap for Processes
	g_risk2processesHashMap.put(p_oSpecificRiskDef, aFuncDefs);       
	for (var j=0; j<aFuncDefs.length; j++) {
		// Func --> ...
		var oFuncDef = aFuncDefs[j];
		// add current process to global set
		// if process was not handled before
        getFurtherObjectsByFunc(oFuncDef);
	}
	
	//update global HashMap for Regulations
	g_risk2regulationsHashMap.put(p_oSpecificRiskDef, aRegulationDefs);
	for (var j=0; j<aRegulationDefs.length; j++) {
		// add current regulations to global set
		// if regulation was not handled before
        getFurtherObjectsByRegulation(aRegulationDefs[j]);
	}	

    //update global HashMap for OrgUnits
    g_risk2orgunitHashMap.put(p_oSpecificRiskDef, aOrgUnitDefs);
    for (var j=0; j<aOrgUnitDefs.length; j++) {
        var oOrgUnitDef = aOrgUnitDefs[j];
        // add current organisational units to global set
		// if organisational unit was not handled before
        getFurtherOrgUnitHierarchyObjectsByOrgUnitDef(oOrgUnitDef);
    } 

    //update global HashMap for AppSysTypes
    g_risk2appSysTypeHashMap.put(p_oSpecificRiskDef, aAppSysTypeDefs);
    for (var j=0; j<aAppSysTypeDefs.length; j++) {
        var oAppSysTypeDef = aAppSysTypeDefs[j];
        // add current application system types to global set
        //if application system was not handled before
        getFurtherObjectsByAppSysType(oAppSysTypeDef);
    }
    
    
    //update global HashMap for Risk Managers
    g_risk2riskManagerHashMap.put(p_oSpecificRiskDef, aAssignedRiskManagers); 
    for (var i=0; i<aAssignedRiskManagers.length ;i++) {g_userGroupSet.add(aAssignedRiskManagers[i]);} // add group to global set
    
    //update global HashMap for Risk Owners
    g_risk2riskOwnerHashMap.put(p_oSpecificRiskDef, aAssignedRiskOwners); 
    for (var i=0; i<aAssignedRiskOwners.length ;i++) {g_userGroupSet.add(aAssignedRiskOwners[i]);} // add group to global set
    
    //update global HashMap for Risk Managers
    g_risk2riskReviewerHashMap.put(p_oSpecificRiskDef, aAssignedRiskReviewers); 
    for (var i=0; i<aAssignedRiskReviewers.length ;i++) {g_userGroupSet.add(aAssignedRiskReviewers[i]);} // add group to global set
    
    
    //update global HashMap for Risk categories
    g_risk2riskCategoryHashMap.put(p_oSpecificRiskDef, aRiskCategoryDefs);
    for (var i=0; i<aRiskCategoryDefs.length; i++) {
        var oRiskCategoryDef = aRiskCategoryDefs[i];
        // add current risk category to global set
        // if risk category was not handled before 
        getFurtherObjectsByRiskCategory(oRiskCategoryDef);
    }
}

/*---------------------------------------------------------------------------------------
    ---- CBA ----
	Determines the objects that are linked to the given control:      
    - Function
    - Technical term
 ---------------------------------------------------------------------------------------*/   
function getFurtherRiskManagementObjectsByControl(p_oControlDef) {
    
    //Control -> Func
    var aFuncDefs = getConnectedObjects(p_oControlDef, Constants.EDGES_OUT, Constants.CT_IS_PERFORMED_AT, Constants.OT_FUNC);
    //update global HashMap for Processes
    g_control2processesHashMap.put(p_oControlDef, aFuncDefs);       
    for (var j=0; j<aFuncDefs.length; j++) {
        // Func --> ...
        var oFuncDef = aFuncDefs[j];
        // add current process to global set  
        getFurtherObjectsByFunc(oFuncDef);
    }
    
    //Control -> TechTerm
    var aTechTermDefs = getConnectedObjects(p_oControlDef, Constants.EDGES_OUT, Constants.CT_AFFECTS, Constants.OT_TECH_TRM);
    //update global HashMap for Regulations
    g_control2regulationsHashMap.put(p_oControlDef, aTechTermDefs);
    for (var j=0; j<aTechTermDefs.length; j++) {
        // add current regulations to global set
        getFurtherObjectsByRegulation(aTechTermDefs[j]);
    }   
}

/*---------------------------------------------------------------------------------------
    Checks for objects of type "User group" if they are risk owners.
    For standard mapping: They are if they have one assigned role object which name starts 
    with "Risk owner".
 ---------------------------------------------------------------------------------------*/        
function isPersonTypeRiskOwner(p_userGroupDef) {
    return hasARCMRole(p_userGroupDef, "riskowner");
}

/*---------------------------------------------------------------------------------------
    Checks for objects of type "User group" if they are risk reviewers.
    For standard mapping: They are if they have one assigned role object which name starts 
    with "Risk reviewer".
 ---------------------------------------------------------------------------------------*/        
function isPersonTypeRiskReviewer(p_userGroupDef) {
    return hasARCMRole(p_userGroupDef, "riskreviewer");
}

/*---------------------------------------------------------------------------------------
    Checks for objects of type "Organisational Unit" if they are part of the orgunit
    hierarchy for the risk management mode. It is if it has one or more assigned risks.
 ---------------------------------------------------------------------------------------*/        
function isOrgunitInRiskManagementOrgUnitHierarchy(p_orgUnitDef) {

    var aRiskDefs = getConnectedObjects(p_orgUnitDef, Constants.EDGES_IN, Constants.CT_AFFECTS, Constants.OT_RISK);     
    return aRiskDefs.length > 0;   
}

/*---------------------------------------------------------------------------------------
    Determines the user groups that are linked to the given hierarchy element:
    - Auditor groups
    - Owner groups
    - Sign-off owner groups
    - Tester groups
    Which of these groups are added to the assignment global maps is controlled by the signature 
    paremters.
 ---------------------------------------------------------------------------------------*/
 function getUserGroupsAssignedToHierarchyElement(  p_oHierarchyElementObjDef
                                                    , p_bGetAuditorGroupsAssignments
                                                    , p_bGetHierarchyOwnerGroupsAssignments
                                                    , p_bGetSignoffOwnerGroupsAssignments
                                                    , p_bGetTesterGroupsAssignments ) {
    //read all assigned usergroups
    var aUserGroups = new Array();
    aUserGroups = aUserGroups.concat( getConnectedObjects(p_oHierarchyElementObjDef, Constants.EDGES_IN, Constants.CT_WRK_IN, Constants.OT_PERS_TYPE) );
    aUserGroups = aUserGroups.concat( getConnectedObjects(p_oHierarchyElementObjDef, Constants.EDGES_IN, Constants.CT_IS_OWN, Constants.OT_PERS_TYPE) );
    
    var aTestAuditorGroups = new Array();
    var aHierarchyOwnerGroups = new Array();
    var aSignoffOwnerGroups = new Array();
    var aTesterGroups = new Array();
    //classify them
    for (var i=0; i<aUserGroups.length; i++) {
        if (p_bGetAuditorGroupsAssignments && isPersonTypeTestAuditor(aUserGroups[i])) {
            aTestAuditorGroups.push(aUserGroups[i]);
			g_userGroupSet.add(aUserGroups[i]);
        }
		if (p_bGetHierarchyOwnerGroupsAssignments && isPersonTypeHierarchyOwner(aUserGroups[i])) {
            aHierarchyOwnerGroups.push(aUserGroups[i]);
			g_userGroupSet.add(aUserGroups[i]);
        }
        if (p_bGetSignoffOwnerGroupsAssignments && isPersonTypeSignOffOwner(aUserGroups[i])) {
            aSignoffOwnerGroups.push(aUserGroups[i]);
			g_userGroupSet.add(aUserGroups[i]);
        }
        if (p_bGetTesterGroupsAssignments && isPersonTypeTester(aUserGroups[i])) {
            aTesterGroups.push(aUserGroups[i]);
			g_userGroupSet.add(aUserGroups[i]);
        }
    }
    //fill the hierarchy2usergroup assignment maps
    g_hierarchy2TestAuditorHashMap.put(p_oHierarchyElementObjDef, aTestAuditorGroups);
	g_hierarchy2HierarchyOwnerHashMap.put(p_oHierarchyElementObjDef, aHierarchyOwnerGroups);
    g_hierarchy2SignoffOwnerHashMap.put(p_oHierarchyElementObjDef, aSignoffOwnerGroups);
    g_hierarchy2TesterGroupHashMap.put(p_oHierarchyElementObjDef, aTesterGroups);
    //fill the testergroup2hierarchy assignmnet map
    for (var i=0; i<aTesterGroups.length; i++) {
        var aHierarchyElements = g_testerGroup2TesterHierarchyHashMap.get(aTesterGroups[i]);
        if (aHierarchyElements == null) {aHierarchyElements = new Array();}
        aHierarchyElements.push(p_oHierarchyElementObjDef);
        g_testerGroup2TesterHierarchyHashMap.put(aTesterGroups[i], aHierarchyElements);
    }
}

/*---------------------------------------------------------------------------------------
    Classify objects of a given model according to risk management modeling
 ---------------------------------------------------------------------------------------*/
function classifyObjectsByRiskManagementModel(p_oModel) {
    
    var aObjectsToClassify = p_oModel.ObjDefList();
    
    //risk diagrams
    if (p_oModel.TypeNum() == Constants.MT_RISK_DGM) {
        for (var i=0; i<aObjectsToClassify.length; i++) {
            //risks -> global risk set
            if (aObjectsToClassify[i].TypeNum() == Constants.OT_RISK) {
                processRiskManagementRisk(aObjectsToClassify[i]);
            }
        }
    }
    	
    //application system type diagrams
    if (p_oModel.TypeNum() == Constants.MT_APPL_SYS_TYPE_DGM) {
        
        for (var i=0; i<aObjectsToClassify.length; i++) {
            //orgunits -> organisational hierarchy
            if (aObjectsToClassify[i].TypeNum() == Constants.OT_APPL_SYS_TYPE) {               
                
                //add only appSysTypes that are linked to an export relevant risk
                var aLinkedRisks = getConnectedObjects(aObjectsToClassify[i], Constants.EDGES_IN, Constants.CT_AFFECTS, Constants.OT_RISK);
                var bExportRelevantRiskLinked = false;
                var bExportRelevantControlLinked = false;
                for (var j=0; j<aLinkedRisks.length && !bExportRelevantRiskLinked; j++) {				
					if (g_approach == "RBA") {
						bExportRelevantRiskLinked = isExportRelevant(aLinkedRisks[j]);
					}
					if (g_approach == "CBA") {
						var aLinkedControls = getConnectedObjects(aLinkedRisks[j], Constants.EDGES_OUT, Constants.CT_IS_REDU_BY, Constants.OT_FUNC);
						for (var k=0; k<aLinkedControls.length && !bExportRelevantControlLinked; k++) {
							bExportRelevantControlLinked = isExportRelevant(aLinkedControls[k]);
						}
					}
                }
                if ( bExportRelevantRiskLinked || bExportRelevantControlLinked ) { 
                    //if application system was not handled before
                    getFurtherObjectsByAppSysType(aObjectsToClassify[i]);
                }
            }
        }    
    }
    
    //organisational charts
    if (p_oModel.TypeNum() == Constants.MT_ORG_CHRT) {
        for (var i=0; i<aObjectsToClassify.length; i++) {
            //person types -> user groups
            if (aObjectsToClassify[i].TypeNum() == Constants.OT_PERS_TYPE) {
                //only "real" usergroups, not their role objects
                if (getRoleObject(aObjectsToClassify[i]) == null) {
                    continue;
                }
                g_userGroupSet.add(aObjectsToClassify[i]);
            }
        }    
    }
    
	getRiskCategoriesFromModel(p_oModel, aObjectsToClassify);
    getOrgUnitsFromModel(p_oModel, aObjectsToClassify);	
	getRegulationsFromModel(p_oModel, aObjectsToClassify);
    getProcessesFromModel(p_oModel, aObjectsToClassify);
    getApplicationSystemTypesFromModel(p_oModel, aObjectsToClassify);
}


/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ~~~~~~~ Restore HashSets and HashMaps from report context if possible ~~~~~~~
                       ~~~~~~~ Survey management ~~~~~~~
  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
  
/**
  * Restore all survey management HashSets and HashMaps
  */
function restoreSurveyManagementFromReportContext() {
    
    if ( Context.getProperty("g_surveyManagementHashSetsFilled") == null 
         || !Context.getProperty("g_surveyManagementHashSetsFilled").equals("true") ) {
         return false;
    }
    
	g_deactivatedLinkSet = restoreStringHashSetFromProperty("g_deactivatedLinkSet", ArisData.getActiveDatabase());
    g_userGroupSet = restoreHashSetFromProperty("g_userGroupSet", ArisData.getActiveDatabase());
    g_userSet = restoreHashSetFromProperty("g_userSet", ArisData.getActiveDatabase());
	
    g_questionnaireTemplateSet = restoreHashSetFromProperty("g_questionnaireTemplateSet", ArisData.getActiveDatabase());
    g_sectionSet = restoreHashSetFromProperty("g_sectionSet", ArisData.getActiveDatabase());
    g_questionSet = restoreHashSetFromProperty("g_questionSet", ArisData.getActiveDatabase());
    g_optionSetSet = restoreHashSetFromProperty("g_optionSetSet", ArisData.getActiveDatabase());
    g_optionSet = restoreHashSetFromProperty("g_optionSet", ArisData.getActiveDatabase());
    g_surveyTaskSet = restoreHashSetFromProperty("g_surveyTaskSet", ArisData.getActiveDatabase()); 
    
    g_questionnaireTemplate2sectionsHashMap = restoreHashMapFromProperty("g_questionnaireTemplate2sectionsHashMap", ArisData.getActiveDatabase());
    g_questionnaireTemplate2surveymanagerHashMap = restoreHashMapFromProperty("g_questionnaireTemplate2surveymanagerHashMap", ArisData.getActiveDatabase());
	g_section2questionnaireTemplateHashMap = restoreHashMapFromProperty("g_section2questionnaireTemplateHashMap", ArisData.getActiveDatabase());
    g_section2subsectionsHashMap = restoreHashMapFromProperty("g_section2subsectionsHashMap", ArisData.getActiveDatabase());
    g_section2topsectionsHashMap = restoreHashMapFromProperty("g_section2topsectionsHashMap", ArisData.getActiveDatabase());
    g_section2questionsHashMap = restoreHashMapFromProperty("g_section2questionsHashMap", ArisData.getActiveDatabase());  
    g_question2sectionsHashMap = restoreHashMapFromProperty("g_question2sectionsHashMap", ArisData.getActiveDatabase());
    g_question2optionsHashMap = restoreHashMapFromProperty("g_question2optionsHashMap", ArisData.getActiveDatabase());
    g_question2optionSetsHashMap = restoreHashMapFromProperty("g_question2optionSetsHashMap", ArisData.getActiveDatabase());
    g_optionSet2optionsHashMap = restoreHashMapFromProperty("g_optionSet2optionsHashMap", ArisData.getActiveDatabase());
	
	g_questionnaireTemplate2surveyTaskHashMap = restoreHashMapFromProperty("g_questionnaireTemplate2surveyTaskHashMap", ArisData.getActiveDatabase());
	g_surveyTask2questionnaireTemplateHashMap = restoreHashMapFromProperty("g_surveyTask2questionnaireTemplateHashMap", ArisData.getActiveDatabase());
	g_surveyTask2questionnaireownerHashMap = restoreHashMapFromProperty("g_surveyTask2questionnaireownerHashMap", ArisData.getActiveDatabase());
	g_surveyTask2surveyreviewerHashMap = restoreHashMapFromProperty("g_surveyTask2surveyreviewerHashMap", ArisData.getActiveDatabase());
	g_surveyTask2surveymanagerHashMap = restoreHashMapFromProperty("g_surveyTask2surveymanagerHashMap", ArisData.getActiveDatabase());  
	g_surveyTask2riskHashMap = restoreHashMapFromProperty("g_surveyTask2riskHashMap", ArisData.getActiveDatabase());
	g_surveyTask2controlHashMap = restoreHashMapFromProperty("g_surveyTask2controlHashMap", ArisData.getActiveDatabase());
	g_surveyTask2testdefinitionHashMap = restoreHashMapFromProperty("g_surveyTask2testdefinitionHashMap", ArisData.getActiveDatabase());
	g_surveyTask2regulationHashMap = restoreHashMapFromProperty("g_surveyTask2regulationHashMap", ArisData.getActiveDatabase());
	g_surveyTask2orgUnitHashMap = restoreHashMapFromProperty("g_surveyTask2orgUnitHashMap", ArisData.getActiveDatabase());
	g_surveyTask2processHashMap = restoreHashMapFromProperty("g_surveyTask2processHashMap", ArisData.getActiveDatabase());
	g_surveyTask2appSysTypeHashMap = restoreHashMapFromProperty("g_surveyTask2appSysTypeHashMap", ArisData.getActiveDatabase());
	g_surveyTask2riskCategoryHashMap = restoreHashMapFromProperty("g_surveyTask2riskCategoryHashMap", ArisData.getActiveDatabase());

	g_question2VisContextQuestionsHashMap = restoreHashMapFromProperty("g_question2VisContextQuestionsHashMap", ArisData.getActiveDatabase());
    g_section2VisContextQuestionsHashMap = restoreHashMapFromProperty("g_section2VisContextQuestionsHashMap", ArisData.getActiveDatabase());

    return true;
}

/**
  * Store all survey management HashSets and HashMaps
  */
function storeSurveyManagementToReportContext() {
    
    if ( !g_surveyManagementHashSetsFilled ) {return;}
    
    Context.setProperty("g_surveyManagementHashSetsFilled", "true");
    
	storeStringHashSetAsProperty("g_deactivatedLinkSet", g_deactivatedLinkSet);
    storeHashSetAsProperty("g_userGroupSet", g_userGroupSet);
    storeHashSetAsProperty("g_userSet", g_userSet);
	
    storeHashSetAsProperty("g_questionnaireTemplateSet", g_questionnaireTemplateSet);
    storeHashSetAsProperty("g_sectionSet", g_sectionSet);
    storeHashSetAsProperty("g_questionSet", g_questionSet);
    storeHashSetAsProperty("g_optionSetSet", g_optionSetSet);
    storeHashSetAsProperty("g_optionSet", g_optionSet);
    storeHashSetAsProperty("g_surveyTaskSet", g_surveyTaskSet);
    
    storeHashMapAsProperty("g_questionnaireTemplate2sectionsHashMap", g_questionnaireTemplate2sectionsHashMap);
    storeHashMapAsProperty("g_questionnaireTemplate2surveymanagerHashMap", g_questionnaireTemplate2surveymanagerHashMap);
	storeHashMapAsProperty("g_section2questionnaireTemplateHashMap", g_section2questionnaireTemplateHashMap);
    storeHashMapAsProperty("g_section2subsectionsHashMap", g_section2subsectionsHashMap);
    storeHashMapAsProperty("g_section2topsectionsHashMap", g_section2topsectionsHashMap);
    storeHashMapAsProperty("g_section2questionsHashMap", g_section2questionsHashMap);
    storeHashMapAsProperty("g_question2sectionsHashMap", g_question2sectionsHashMap);
    storeHashMapAsProperty("g_question2optionsHashMap", g_question2optionsHashMap);
    storeHashMapAsProperty("g_question2optionSetsHashMap", g_question2optionSetsHashMap);
    storeHashMapAsProperty("g_optionSet2optionsHashMap", g_optionSet2optionsHashMap);
	
	storeHashMapAsProperty("g_questionnaireTemplate2surveyTaskHashMap", g_questionnaireTemplate2surveyTaskHashMap);
	storeHashMapAsProperty("g_surveyTask2questionnaireTemplateHashMap", g_surveyTask2questionnaireTemplateHashMap);
	storeHashMapAsProperty("g_surveyTask2questionnaireownerHashMap", g_surveyTask2questionnaireownerHashMap);
	storeHashMapAsProperty("g_surveyTask2surveyreviewerHashMap", g_surveyTask2surveyreviewerHashMap);
	storeHashMapAsProperty("g_surveyTask2surveymanagerHashMap", g_surveyTask2surveymanagerHashMap);  
    storeHashMapAsProperty("g_surveyTask2riskHashMap", g_surveyTask2riskHashMap);
	storeHashMapAsProperty("g_surveyTask2riskHashMap", g_surveyTask2riskHashMap);
	storeHashMapAsProperty("g_surveyTask2controlHashMap", g_surveyTask2controlHashMap);
	storeHashMapAsProperty("g_surveyTask2testdefinitionHashMap", g_surveyTask2testdefinitionHashMap);
	storeHashMapAsProperty("g_surveyTask2regulationHashMap", g_surveyTask2regulationHashMap);
	storeHashMapAsProperty("g_surveyTask2orgUnitHashMap", g_surveyTask2orgUnitHashMap);
	storeHashMapAsProperty("g_surveyTask2processHashMap", g_surveyTask2processHashMap);
	storeHashMapAsProperty("g_surveyTask2appSysTypeHashMap", g_surveyTask2appSysTypeHashMap);
	storeHashMapAsProperty("g_surveyTask2riskCategoryHashMap", g_surveyTask2riskCategoryHashMap);
	
	storeHashMapAsProperty("g_question2VisContextQuestionsHashMap", g_question2VisContextQuestionsHashMap);
    storeHashMapAsProperty("g_section2VisContextQuestionsHashMap", g_section2VisContextQuestionsHashMap);
}

/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    ~~~~~~~ Determine start objects ~~~~~~~
                        ~~~~~~~ Survey management ~~~~~~~
  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
/*---------------------------------------------------------------------------------------
    Determine the start objects relevant for basic ARCM survey management depending on 
    the selected context: 
    - Group: 
    - Model: survey management diagrams, questionnaire template allocations
    - ObjDef: Questionnaire template
	Processes the found questionnaire templates only if the global flag 
	"g_bSurveyManagementExport" is true (see determineMode() in arcm-common.js).
    Fills the corresponding global HashSets.
    Returns an array of all found relevant objects.
 ---------------------------------------------------------------------------------------*/   
function getSurveyManagementSelectionByContext() {
    
    Context.writeStatus("Searching survey management data");
    //start object search only if the object GUIDs are not already passed as Context properties
    if (!restoreSurveyManagementFromReportContext()) {
        buildAllowedModels();        
		// context "Group"
		var oSelectedGroups = ArisData.getSelectedGroups();
        if(oSelectedGroups.length == 0){
            var databases = ArisData.getSelectedDatabases();
            if (databases.length != 0) {
                oSelectedGroups = [databases[0].RootGroup()];
            }
        }
		if (oSelectedGroups.length > 0) {
			for (var i = 0; i < oSelectedGroups.length; i++ ) {
				getSurveyManagementRelevantObjects(oSelectedGroups[i]);
			}
		}
		
		// context "Model"    
		var oSelectedModels = ArisData.getSelectedModels();
		if (oSelectedModels.length > 0) {
			for (var i = 0; i < oSelectedModels.length; i++ ) {
				
				//model must be marked as export relevant
				if (isExportRelevant(oSelectedModels[i])) {
					classifyObjectsBySurveyManagementModel(oSelectedModels[i]);
				}
			}
		}
		
		// context "ObjDef"    
		var oSelectedObjDefs = ArisData.getSelectedObjDefs();
        if (oSelectedObjDefs.length > 0) {
            for (var i = 0; i < oSelectedObjDefs.length; i++ ) {       
                if (oSelectedObjDefs[i].TypeNum() == Constants.OT_SURVEY_QUEST_TMPL) {
                    processSurveyManagementQuestionnaireTemplate(oSelectedObjDefs[i]);
                }
            }
		}
        
		//after global HashSets are filled:
		//Find all users to all found user groups   
		classifyUsers();
        removeEntriesFromUrgUnitSetIfinTesterSet();
		        
		g_surveyManagementHashSetsFilled = true;
        storeSurveyManagementToReportContext();
    }
}

/*---------------------------------------------------------------------------------------
    Determine the relevant objects in the given group
---------------------------------------------------------------------------------------*/
function getSurveyManagementRelevantObjects(p_oGroup) {
    
    // Get export relevant questionnaire templates objects on group
    var oQuestionnaireTemplateDefs = p_oGroup.ObjDefListFilter(Constants.OT_SURVEY_QUEST_TMPL);
    for (var i = 0; i < oQuestionnaireTemplateDefs.length; i++ ) {                
        var oQuestionnaireTemplateDef = oQuestionnaireTemplateDefs[i];
        processSurveyManagementQuestionnaireTemplate(oQuestionnaireTemplateDef);
    }
    
    // Get objects in export relevant models
    var oModels = p_oGroup.ModelList();
    for (var i = 0; i < oModels.length; i++ ) {

        //model must be marked as export relevant        
        if (isExportRelevant(oModels[i])) {
            classifyObjectsBySurveyManagementModel(oModels[i]);
        }
    }
    
    //Search by recursion in child groups
    var oChildGroups = p_oGroup.Childs();
    for (var i = 0; i < oChildGroups.length; i++ ) {
        getSurveyManagementRelevantObjects(oChildGroups[i]);
    }
    
}

/*---------------------------------------------------------------------------------------
    Determines the objects that are linked to the given questionnaire template:
    - Section
 ---------------------------------------------------------------------------------------*/   
function processSurveyManagementQuestionnaireTemplate(p_oQuestionnaireTemplateDef) {

	//if the global flag "g_bSurveyManagementExport" is not set then don't process anything
	if (!g_bSurveyManagementExport) {
		return;
	}

    //questionnaire template must be marked as export relevant
    if (isExportRelevant(p_oQuestionnaireTemplateDef)) {
    
		//if questionnaire template was handled before then skip it, else process it further
		if ( !g_questionnaireTemplateSet.add(p_oQuestionnaireTemplateDef) ) {
			return;    
		}
    
        // Questionnaire Template --> Section
        var aSectionDefs = getConnectedObjects(p_oQuestionnaireTemplateDef, Constants.EDGES_OUT, Constants.CT_SURVEY_CONTAINS, Constants.OT_SURVEY_SECTION);
        
		// Questionnaire Template --> Survey Task
		var aSurveyTaskDefs = getConnectedObjects(p_oQuestionnaireTemplateDef, Constants.EDGES_OUT, Constants.CT_IS_INITIATED_BY, Constants.OT_SURVEY_TASK);
		
		// Questionnaire Template --> Survey managers
        var aPersTypeDefs = getConnectedObjects(p_oQuestionnaireTemplateDef, Constants.EDGES_IN, Constants.CT_IS_ASSIG_1, Constants.OT_PERS_TYPE);
        var aSurveyManagers = new Array();
        for (var j=0; j<aPersTypeDefs.length; j++) {
            if (isPersonTypeSurveyManager(aPersTypeDefs[j]))  {aSurveyManagers.push(aPersTypeDefs[j]);}
        } 
		
        //update global HashMap for Sections
        g_questionnaireTemplate2sectionsHashMap.put(p_oQuestionnaireTemplateDef, aSectionDefs);
        for (var j=0; j<aSectionDefs.length; j++) {       
            var oSectionDef = aSectionDefs[j];

            var aTemplates = g_section2questionnaireTemplateHashMap.get(oSectionDef);
            if (aTemplates == null) {aTemplates = new Array();}
            aTemplates.push(p_oQuestionnaireTemplateDef);
            g_section2questionnaireTemplateHashMap.put(oSectionDef, aTemplates);
            
            // Section --> ...        
            //mark section as explicitly positioned within this questionnaire template
            markSectionAsExplicitForQT(oSectionDef, p_oQuestionnaireTemplateDef);
            
            // add current section to global set
            g_sectionSet.add(oSectionDef);
            
            if (!isSectionProcessedAlreadyForThisQT(p_oQuestionnaireTemplateDef, oSectionDef)) {
                getFurtherObjectsBySection(oSectionDef, null, p_oQuestionnaireTemplateDef, false);
            }
        }
		
		//update global HashMap for Survey Tasks
		g_questionnaireTemplate2surveyTaskHashMap.put(p_oQuestionnaireTemplateDef, aSurveyTaskDefs);
		for (var k=0; k<aSurveyTaskDefs.length; k++) {
			var oSurveyTaskDef = aSurveyTaskDefs[k];
			
			var aTemplates = g_surveyTask2questionnaireTemplateHashMap.get(oSurveyTaskDef);
            if (aTemplates == null) {aTemplates = new Array();}
            aTemplates.push(p_oQuestionnaireTemplateDef);
            g_surveyTask2questionnaireTemplateHashMap.put(oSurveyTaskDef, aTemplates);
			
			// Survey Task --> ...
            // add current survey task to global set  
            if (g_surveyTaskSet.add(oSurveyTaskDef)) {               
                getFurtherObjectsBySurveyTask(oSurveyTaskDef);
            }
		}
		
		//update global HashMap for Survey Managers
        g_questionnaireTemplate2surveymanagerHashMap.put(p_oQuestionnaireTemplateDef, aSurveyManagers); 
        for (var i=0; i<aSurveyManagers.length ;i++) {g_userGroupSet.add(aSurveyManagers[i]);} // add group to global set
    }
}

/*---------------------------------------------------------------------------------------
    Determines the objects that are linked to the given section:
    - Section
    - Question
 ---------------------------------------------------------------------------------------*/
function getFurtherObjectsBySection(p_oSectionDef, p_oImplicitSectionPlacingSectionDef, p_oQuestionnaireTemplateDef, p_bImplicit) {  
    
    // Section --> Section
    var aSubSectionDefs = getConnectedObjects(p_oSectionDef, Constants.EDGES_OUT, Constants.CT_SURVEY_CONTAINS, Constants.OT_SURVEY_SECTION);  
    
    //update global HashMap for Sections
    g_section2subsectionsHashMap.put(p_oSectionDef, aSubSectionDefs);
    
    for (var j=0; j<aSubSectionDefs.length; j++) {
        
        var oSubSectionDef = aSubSectionDefs[j];
        //update global HashMap for Section
        var aTopSections = g_section2topsectionsHashMap.get(oSubSectionDef);
        if (aTopSections == null) {aTopSections = new Array();}
        aTopSections.push(p_oSectionDef);
        g_section2topsectionsHashMap.put(oSubSectionDef, aTopSections);
        
        // Section --> ...        
        //mark section as explicitly positioned within this questionnaire template
        if (!p_bImplicit) {
            markSectionAsExplicitForQT(oSubSectionDef, p_oQuestionnaireTemplateDef);
        }
         
        // add current section to global set
        g_sectionSet.add(oSubSectionDef);
        
        if (!isSectionProcessedAlreadyForThisQT(p_oQuestionnaireTemplateDef, oSubSectionDef)) { 
           if (p_bImplicit) {
                getFurtherObjectsBySection(oSubSectionDef, p_oImplicitSectionPlacingSectionDef, p_oQuestionnaireTemplateDef, p_bImplicit);
            }
            else {
                getFurtherObjectsBySection(oSubSectionDef, p_oSectionDef, p_oQuestionnaireTemplateDef, p_bImplicit);
            }
        }
    }
    
    // Section --> Question
    var aQuestionDefs = getConnectedObjects(p_oSectionDef, Constants.EDGES_OUT, Constants.CT_SURVEY_CONTAINS, Constants.OT_SURVEY_QUESTION);  
    
    //update global HashMap for Section->Questions
    g_section2questionsHashMap.put(p_oSectionDef, aQuestionDefs);
    //update global HashMap for Question->Sections
    for (var m=0; m<aQuestionDefs.length; m++) {
        var aQuestion2Sections = g_question2sectionsHashMap.get(aQuestionDefs[m]);
        if (aQuestion2Sections == null) {
            aQuestion2Sections = new Array();
            g_question2sectionsHashMap.put(aQuestionDefs[m], aQuestion2Sections);
        }
        aQuestion2Sections.push(p_oSectionDef);
    }   
     
    for (var j=0; j<aQuestionDefs.length; j++) {
        
        // Question --> ...
        var oQuestionDef = aQuestionDefs[j];  
            
        // add current question to global set
        g_questionSet.add(oQuestionDef);
        
        //mark question as explicitly positioned within this questionnaire template
        if (!p_bImplicit) {
            markQuestionAsExplicitForQT(oQuestionDef, p_oQuestionnaireTemplateDef);
        }
        
        if (!isQuestionProcessedAlreadyForThisQTandSection(p_oQuestionnaireTemplateDef, p_oSectionDef, oQuestionDef)) {
            getFurtherObjectsByQuestion(oQuestionDef, p_oSectionDef, p_oImplicitSectionPlacingSectionDef, p_oQuestionnaireTemplateDef, p_bImplicit);
        }
    }
}

/*---------------------------------------------------------------------------------------
    Determines the objects that are linked to the given question:
    - Option
    - OptionSet
 ---------------------------------------------------------------------------------------*/
function getFurtherObjectsByQuestion(p_oQuestionDef, p_oImplicitQuestionPlacingSectionDef, p_oImplicitSectionPlacingSectionDef, p_oQuestionnaireTemplateDef, p_bImplicit) {
    
    // Question --> Option
    var aOptionDefs = getConnectedObjects(p_oQuestionDef, Constants.EDGES_OUT, Constants.CT_SURVEY_HAS, Constants.OT_SURVEY_OPTION);  
    
    //update global HashMap for Options
    g_question2optionsHashMap.put(p_oQuestionDef, aOptionDefs);       
    for (var j=0; j<aOptionDefs.length; j++) {
        g_optionSet.add(aOptionDefs[j]);
        getFurtherObjectsByOption(aOptionDefs[j], p_oQuestionDef, p_oImplicitQuestionPlacingSectionDef, p_oImplicitSectionPlacingSectionDef, p_oQuestionnaireTemplateDef, p_bImplicit);
    }

    // Question --> OptionSet
    var aOptionSetDefs = getConnectedObjects(p_oQuestionDef, Constants.EDGES_OUT, Constants.CT_SURVEY_CONTAINS, Constants.OT_SURVEY_OPTION_SET);  
    
    //update global HashMap for OptionSets
    g_question2optionSetsHashMap.put(p_oQuestionDef, aOptionSetDefs);       
    for (var j=0; j<aOptionSetDefs.length; j++) {
        // Question --> ...
        var oOptionSetDef = aOptionSetDefs[j];
        // add current question to global set
        g_optionSetSet.add(oOptionSetDef);                              
        getFurtherObjectsByOptionSet(oOptionSetDef, p_oQuestionDef, p_oImplicitQuestionPlacingSectionDef, p_oImplicitSectionPlacingSectionDef, p_oQuestionnaireTemplateDef, p_bImplicit);
    }
}

/*---------------------------------------------------------------------------------------
    Determines the objects that are linked to the given option set:
    - Option
 ---------------------------------------------------------------------------------------*/
function getFurtherObjectsByOptionSet(p_oOptionSetDef, p_oActivatingQuestionObjDef, p_oImplicitQuestionPlacingSectionDef, p_oImplicitSectionPlacingSectionDef, p_oQuestionnaireTemplateDef, p_bImplicit) {

    // OptionSet --> Option
    var aOptionDefs = getConnectedObjects(p_oOptionSetDef, Constants.EDGES_OUT, Constants.CT_SURVEY_HAS, Constants.OT_SURVEY_OPTION);  
    
    //update global HashMap for Options
    g_optionSet2optionsHashMap.put(p_oOptionSetDef, aOptionDefs);
    for (var j=0; j<aOptionDefs.length; j++) {
        g_optionSet.add(aOptionDefs[j]);
        getFurtherObjectsByOption(aOptionDefs[j], p_oActivatingQuestionObjDef, p_oImplicitQuestionPlacingSectionDef, p_oImplicitSectionPlacingSectionDef, p_oQuestionnaireTemplateDef, p_bImplicit);
    }
    
}

/*---------------------------------------------------------------------------------------
    Determines the objects that are linked to the given option:
    - dependent visible Questions
    - dependent visible Sections
 ---------------------------------------------------------------------------------------*/
function getFurtherObjectsByOption(p_oOptionDef, p_oActivatingQuestionObjDef, p_oImplicitQuestionPlacingSectionDef, p_oImplicitSectionPlacingSectionDef, p_oQuestionnaireTemplateDef, p_bImplicit) {
    
    //we support implicit sections and questions only one level deep i.e. implicit sections/questions cannot have further implicit sections or questions - all such relations are ignored
    if (p_bImplicit) {
        return;
    }
    
    //we support implicit sections and questions only at single choice and multi choice questions (i.e. where the interviewee has to select options, not the reviewer)
    var oQuestionTypeAttribute = p_oActivatingQuestionObjDef.Attribute( Constants.AT_QUESTION_TYPE, g_nLoc);
    var iUnitTypeNum = oQuestionTypeAttribute.MeasureUnitTypeNum();
    if (iUnitTypeNum != Constants.AVT_SINGLE_CHOICE && iUnitTypeNum != Constants.AVT_MULTIPLE_CHOICE ) {
        return;
    }
    
    // Option --> visibility dependent Question
    var aVisDepQuestionDefs = getConnectedObjects(p_oOptionDef, Constants.EDGES_OUT, Constants.CT_ACTIVATES_1, Constants.OT_SURVEY_QUESTION);
    
    for (var i=0; i<aVisDepQuestionDefs.length; i++) {
        
        // Question --> visibility context Question
        // ask the cache first if we did this look-up already before
        var aContextQuestionDefs = g_question2VisContextQuestionsHashMap.get(aVisDepQuestionDefs[i]);
        if (aContextQuestionDefs == null) {
            aContextQuestionDefs = getConnectedObjects(aVisDepQuestionDefs[i], Constants.EDGES_OUT, Constants.CT_IS_VALID_CONTEXT_OF, Constants.OT_SURVEY_QUESTION);
            g_question2VisContextQuestionsHashMap.put(aVisDepQuestionDefs[i], aContextQuestionDefs);
        }
        
        //if there are context questions then at least one of them must be the activating question, otherwise this is no possible implicit placing
        var bContextFulfilled = aContextQuestionDefs.length == 0 || isObjDefContained(p_oActivatingQuestionObjDef, aContextQuestionDefs);
        
        //store the info for a possible implicit placing
        if (bContextFulfilled) {
            var impPlac = new implicitPlacing(aVisDepQuestionDefs[i], p_oImplicitQuestionPlacingSectionDef, p_oImplicitQuestionPlacingSectionDef, p_oActivatingQuestionObjDef, aContextQuestionDefs);   
            var a_implicitQuestionPlacings = g_possibleImplicitQuestions.get(p_oQuestionnaireTemplateDef);
            if (a_implicitQuestionPlacings == null) {
                a_implicitQuestionPlacings = new Array();
                 g_possibleImplicitQuestions.put(p_oQuestionnaireTemplateDef, a_implicitQuestionPlacings);
            }
            a_implicitQuestionPlacings.push(impPlac);
        }
        
        // Question --> ...
        getFurtherObjectsByQuestion(aVisDepQuestionDefs[i], p_oImplicitQuestionPlacingSectionDef, p_oImplicitQuestionPlacingSectionDef, p_oImplicitSectionPlacingSectionDef, p_oQuestionnaireTemplateDef, true);
    }
    
    // Option --> visibility dependent Section
    var aVisDepSectionDefs = getConnectedObjects(p_oOptionDef, Constants.EDGES_OUT, Constants.CT_ACTIVATES_1, Constants.OT_SURVEY_SECTION);
     
    for (var j=0; j<aVisDepSectionDefs.length; j++) {
        
        // Section --> visibility context Question
        // ask the cache first if we did this look-up already before
        var aContextQuestionDefs = g_section2VisContextQuestionsHashMap.get(aVisDepSectionDefs[j]);
        if (aContextQuestionDefs == null) {
            aContextQuestionDefs = getConnectedObjects(aVisDepSectionDefs[j], Constants.EDGES_OUT, Constants.CT_IS_VALID_CONTEXT_OF, Constants.OT_SURVEY_QUESTION);
            g_section2VisContextQuestionsHashMap.put(aVisDepSectionDefs[j], aContextQuestionDefs);
        }
        
        //if there are context questions then at least one of them must be the activating question, otherwise this is no possible implicit placing
        var bContextFulfilled = aContextQuestionDefs.length == 0 || isObjDefContained(p_oActivatingQuestionObjDef, aContextQuestionDefs);
        
        //store the info for a possible implicit placing
        if (bContextFulfilled) {
            if (p_oImplicitSectionPlacingSectionDef != null) {
                //placing as sub section under a section
                var impPlac = new implicitPlacing(aVisDepSectionDefs[j], p_oImplicitSectionPlacingSectionDef, p_oImplicitQuestionPlacingSectionDef, p_oActivatingQuestionObjDef, aContextQuestionDefs);   
                var a_implicitSectionPlacings = g_possibleImplicitSections.get(p_oQuestionnaireTemplateDef);
                if (a_implicitSectionPlacings == null) {
                    a_implicitSectionPlacings = new Array();
                     g_possibleImplicitSections.put(p_oQuestionnaireTemplateDef, a_implicitSectionPlacings);
                }
                a_implicitSectionPlacings.push(impPlac);
            } else {
                //placing as top level section
                var impPlac = new implicitPlacing(aVisDepSectionDefs[j], null, p_oImplicitQuestionPlacingSectionDef, p_oActivatingQuestionObjDef, aContextQuestionDefs);   
                var a_implicitSectionPlacings = g_possibleImplicitTopLevelSections.get(p_oQuestionnaireTemplateDef);
                if (a_implicitSectionPlacings == null) {
                    a_implicitSectionPlacings = new Array();
                    g_possibleImplicitTopLevelSections.put(p_oQuestionnaireTemplateDef, a_implicitSectionPlacings);
                }
                a_implicitSectionPlacings.push(impPlac);
            }
        }
        
        // Section --> ...
        g_sectionSet.add(aVisDepSectionDefs[j]);
        getFurtherObjectsBySection(aVisDepSectionDefs[j], p_oImplicitSectionPlacingSectionDef, p_oQuestionnaireTemplateDef, true);
    }

}

/*---------------------------------------------------------------------------------------
    Determines the objects that are linked to the given survey task:
    - Person type
	- Risk (only if test management or risk management shall be classified)
	- Control (only if test management shall be classified)
	- Testdefinition (only if test management shall be classified)
	- Application system type
	- Organisational unit
	- Function
	- Technical term
	- Risk category
 ---------------------------------------------------------------------------------------*/
function getFurtherObjectsBySurveyTask(p_oSurveyTaskDef) {

	// Survey Task --> PersType        
	var aPersTypeDefs = getConnectedObjects(p_oSurveyTaskDef, Constants.EDGES_IN, Constants.CT_IS_ASSIG_1, Constants.OT_PERS_TYPE);
	//differentiate the roles
	var aQuestionnaireOwners = new Array();
	var aSurveyReviewers = new Array();
	var aSurveyManagers = new Array();
	for (var j=0; j<aPersTypeDefs.length; j++) {
		if (isPersonTypeQuestionnaireOwner(aPersTypeDefs[j])) {aQuestionnaireOwners.push(aPersTypeDefs[j]);}
		if (isPersonTypeSurveyReviewer(aPersTypeDefs[j])) {aSurveyReviewers.push(aPersTypeDefs[j]);}
		if (isPersonTypeSurveyManager(aPersTypeDefs[j])) {aSurveyManagers.push(aPersTypeDefs[j]);}
	}
	//update global HashMap for Questionnaire Owners
	g_surveyTask2questionnaireownerHashMap.put(p_oSurveyTaskDef, aQuestionnaireOwners); 
	for (var i=0; i<aQuestionnaireOwners.length ;i++) {g_userGroupSet.add(aQuestionnaireOwners[i]);} // add group to global set
	//update global HashMap for Survey Reviewers
	g_surveyTask2surveyreviewerHashMap.put(p_oSurveyTaskDef, aSurveyReviewers); 
	for (var i=0; i<aSurveyReviewers.length ;i++) {g_userGroupSet.add(aSurveyReviewers[i]);} // add group to global set
	//update global HashMap for Survey Managers
	g_surveyTask2surveymanagerHashMap.put(p_oSurveyTaskDef, aSurveyManagers); 
	for (var i=0; i<aSurveyManagers.length ;i++) {g_userGroupSet.add(aSurveyManagers[i]);} // add group to global set

	// Survey Task --> Risk   
	var aRiskDefs = getConnectedObjects(p_oSurveyTaskDef, Constants.EDGES_OUT, Constants.CT_CONC_3, Constants.OT_RISK);	
	//update global HashMap for Risks - only if g_bTestManagementExport or g_bRiskManagementExport is true
	if (g_bTestManagementExport || g_bRiskManagementExport) {
		g_surveyTask2riskHashMap.put(p_oSurveyTaskDef, aRiskDefs);
		// Risk --> ...
		for (var j=0; j<aRiskDefs.length; j++) {
			var oRiskDef = aRiskDefs[j];
			if (g_riskSet.add(oRiskDef)) {
				if (g_bTestManagementExport) {
					getFurtherTestManagementObjectsByRisk(oRiskDef);
				}
				if (g_bRiskManagementExport) {
					getFurtherRiskManagementObjectsByRisk(oRiskDef);
				}
			}
		}
	}
	
	// Survey Task --> Control
    var aCtrlDefs = getConnectedObjects(p_oSurveyTaskDef, Constants.EDGES_OUT, Constants.CT_CONC_3, Constants.OT_FUNC, Constants.ST_CONTR);
	//update global HashMap for Controls - only if g_bTestManagementExport is true
	if (g_bTestManagementExport) {
		g_surveyTask2controlHashMap.put(p_oSurveyTaskDef, aCtrlDefs);
		// Control --> ...
		for (var j=0; j<aCtrlDefs.length; j++) {
			var oCtrlDef = aCtrlDefs[j];
			if (g_controlSet.add(oCtrlDef)) {
				getFurtherTestManagementObjectsByControl(oCtrlDef);
			}
		}
	}
	
	// Survey Task --> Testdefinition
	var aTestDefs = getConnectedObjects(p_oSurveyTaskDef, Constants.EDGES_OUT, Constants.CT_CONC_3, Constants.OT_TEST_DEFINITION);
	//update global HashMap for Testdefinitions - only if g_bTestManagementExport is true
	if (g_bTestManagementExport) {
		g_surveyTask2testdefinitionHashMap.put(p_oSurveyTaskDef, aTestDefs);
		// Testdefinition --> ...
		for (var j=0; j<aTestDefs.length; j++) {
			var oTestDef = aTestDefs[j];
			if (g_testdefinitionSet.add(oTestDef)) {
				getFurtherObjectsByTest(oTestDef);
			}
		}
	}
	
	// Survey Task --> Application system type
	var aAppSysTypeDefs = getConnectedObjects(p_oSurveyTaskDef, Constants.EDGES_OUT, Constants.CT_CONC_3, Constants.OT_APPL_SYS_TYPE);
	//update global HashMap for AppSysTypes
    g_surveyTask2appSysTypeHashMap.put(p_oSurveyTaskDef, aAppSysTypeDefs);
	// Application system type --> ...
    for (var j=0; j<aAppSysTypeDefs.length; j++) {
        var oAppSysTypeDef = aAppSysTypeDefs[j];
        getFurtherObjectsByAppSysType(oAppSysTypeDef);
    }
	
	// Survey Task --> Organisatonal unit
	var aOrgUnitDefs = getConnectedObjects(p_oSurveyTaskDef, Constants.EDGES_OUT, Constants.CT_CONC_3, Constants.OT_ORG_UNIT);
	//update global HashMap for OrgUnits
    g_surveyTask2orgUnitHashMap.put(p_oSurveyTaskDef, aOrgUnitDefs);
	// Organisational unit --> ...
    for (var i=0; i<aOrgUnitDefs.length ;i++) {
        getFurtherOrgUnitHierarchyObjectsByOrgUnitDef(aOrgUnitDefs[i]);
    }
	
	// Survey Task --> Function
    var aFuncDefs = new Array();
	aFuncDefs = aFuncDefs.concat( getConnectedObjects(p_oSurveyTaskDef, Constants.EDGES_OUT, Constants.CT_CONC_3, Constants.OT_FUNC, Constants.ST_FUNC) );
    aFuncDefs = aFuncDefs.concat( getConnectedObjects(p_oSurveyTaskDef, Constants.EDGES_OUT, Constants.CT_CONC_3, Constants.OT_FUNC, Constants.ST_VAL_ADD_CHN_SML_1) );
    aFuncDefs = aFuncDefs.concat( getConnectedObjects(p_oSurveyTaskDef, Constants.EDGES_OUT, Constants.CT_CONC_3, Constants.OT_FUNC, Constants.ST_VAL_ADD_CHN_SML_2) ); 
    aFuncDefs = aFuncDefs.concat( getConnectedObjects(p_oSurveyTaskDef, Constants.EDGES_OUT, Constants.CT_CONC_3, Constants.OT_FUNC, "b9f9ecb1-45e9-11e7-43b7-08002721906d") ); //processing activity

	//update global HashMap for Processes
    g_surveyTask2processHashMap.put(p_oSurveyTaskDef, aFuncDefs);
    // Function --> ...
	for (var i = 0; i < aFuncDefs.length; i++ ) {
        getFurtherObjectsByFunc(aFuncDefs[i]);
    }
	
	// Survey Task --> Technical term
	var aTechTermDefs = getConnectedObjects(p_oSurveyTaskDef, Constants.EDGES_OUT, Constants.CT_CONC_3, Constants.OT_TECH_TRM);
	var aRegulationDefs = new Array();
	for (var i=0; i<aTechTermDefs.length; i++) {
		if (isRegulationItem(aTechTermDefs[i])) {
			aRegulationDefs.push(aTechTermDefs[i]);
		}
	}
	//update global HashMap for Regulations
    g_surveyTask2regulationHashMap.put(p_oSurveyTaskDef, aRegulationDefs);  
    // Technical term -> ...
	for (var i=0; i<aRegulationDefs.length; i++) {
        getFurtherObjectsByRegulation(aRegulationDefs[i]);
    }
	
	// Survey Task --> Risk category
	var aRiskCategoryDefs = getConnectedObjects(p_oSurveyTaskDef, Constants.EDGES_OUT, Constants.CT_CONC_3, Constants.OT_RISK_CATEGORY);
	//update global HashMap for Risk categories
    g_surveyTask2riskCategoryHashMap.put(p_oSurveyTaskDef, aRiskCategoryDefs);
	// Risk category -> ...
    for (var i=0; i<aRiskCategoryDefs.length; i++) {
        getFurtherObjectsByRiskCategory(aRiskCategoryDefs[i]);
    }	
}

/*---------------------------------------------------------------------------------------
    Helper method to mark the given section as explicitly positioned within the given
    questionnaire template.
 ---------------------------------------------------------------------------------------*/
function markSectionAsExplicitForQT(p_oSectionDef, p_oQuestionnaireTemplateDef) {
    var a_explicitSections = g_explicitPositionSections.get(p_oQuestionnaireTemplateDef);
    if (a_explicitSections == null) {
        a_explicitSections = new Array();
        g_explicitPositionSections.put(p_oQuestionnaireTemplateDef, a_explicitSections);
    }
    a_explicitSections.push(p_oSectionDef);
}

/*---------------------------------------------------------------------------------------
    Helper method to mark the given question as explicitly positioned within the given
    questionnaire template.
 ---------------------------------------------------------------------------------------*/
function markQuestionAsExplicitForQT(p_oQuestionDef, p_oQuestionnaireTemplateDef) {
    var a_explicitQuestions = g_explicitPositionQuestions.get(p_oQuestionnaireTemplateDef);
    if (a_explicitQuestions == null) {
        a_explicitQuestions = new Array();
        g_explicitPositionQuestions.put(p_oQuestionnaireTemplateDef, a_explicitQuestions);
    }
    a_explicitQuestions.push(p_oQuestionDef);
}

/*---------------------------------------------------------------------------------------
    Helper method to fill the recursion stop HashSet for sections
 ---------------------------------------------------------------------------------------*/
function isSectionProcessedAlreadyForThisQT(p_oQuestionnaireTemplateDef, p_oSectionDef) {
    return !g_sectionRecursionStop.add(p_oQuestionnaireTemplateDef.GUID() + "|" + p_oSectionDef.GUID());
}

/*---------------------------------------------------------------------------------------
    Helper method to fill the recursion stop HashSet for questions
 ---------------------------------------------------------------------------------------*/
function isQuestionProcessedAlreadyForThisQTandSection(p_oQuestionnaireTemplateDef, p_oSectionDef, p_oQuestionDef) {
    return !g_questionRecursionStop.add(p_oQuestionnaireTemplateDef.GUID() + "|" + p_oSectionDef.GUID() + "|" + p_oQuestionDef.GUID());
}

/*---------------------------------------------------------------------------------------
    Helper method to fill the recursion stop HashSet for sections
 ---------------------------------------------------------------------------------------*/
function isOptionSetProcessedAlreadyForThisQT(p_oQuestionnaireTemplateDef, p_oOptionSetDef) {
    return !g_optionSetRecursionStop.add(p_oQuestionnaireTemplateDef.GUID() + "|" + p_oOptionSetDef.GUID());
}

/*---------------------------------------------------------------------------------------
    Helper method to fill the recursion stop HashSet for sections
 ---------------------------------------------------------------------------------------*/
function isOptionProcessedAlreadyForThisQT(p_oQuestionnaireTemplateDef, p_oOptionDef) {
    return !g_optionRecursionStop.add(p_oQuestionnaireTemplateDef.GUID() + "|" + p_oOptionDef.GUID());
}

/*---------------------------------------------------------------------------------------
    Checks for objects of type "User group" if they are questionnaire owners.
    For standard mapping: They are if they have one assigned role object which name starts 
    with "Interviewee".
 ---------------------------------------------------------------------------------------*/        
function isPersonTypeQuestionnaireOwner(p_userGroupDef) {
    return hasARCMRole(p_userGroupDef, "questionnaireowner");
}

/*---------------------------------------------------------------------------------------
    Checks for objects of type "User group" if they are survey reviewers.
    For standard mapping: They are if they have one assigned role object which name starts 
    with "Survey reviewer".
 ---------------------------------------------------------------------------------------*/        
function isPersonTypeSurveyReviewer(p_userGroupDef) {
    return hasARCMRole(p_userGroupDef, "surveyreviewer");
}

/*---------------------------------------------------------------------------------------
    Checks for objects of type "User group" if they are survey managers.
    For standard mapping: They are if they have one assigned role object which name starts 
    with "Survey manager".
 ---------------------------------------------------------------------------------------*/        
function isPersonTypeSurveyManager(p_userGroupDef) {
    return hasARCMRole(p_userGroupDef, "surveymanager");
}

/*---------------------------------------------------------------------------------------
    Classify objects of a given model according to survey management modeling
 ---------------------------------------------------------------------------------------*/
function classifyObjectsBySurveyManagementModel(p_oModel) {

    var aObjectsToClassify = p_oModel.ObjDefList();
    
    //survey management diagrams
    if (p_oModel.TypeNum() == Constants.MT_SURVEY_MGMT) {
        for (var i=0; i<aObjectsToClassify.length; i++) {
            //questionnaire templates in this model
            if (aObjectsToClassify[i].TypeNum() == Constants.OT_SURVEY_QUEST_TMPL) {
                processSurveyManagementQuestionnaireTemplate(aObjectsToClassify[i]);
            }
        }
    }
	getRiskCategoriesFromModel(p_oModel, aObjectsToClassify);    
	getRegulationsFromModel(p_oModel, aObjectsToClassify);	
	getOrgUnitsFromModel(p_oModel, aObjectsToClassify);        
    getProcessesFromModel(p_oModel, aObjectsToClassify);
    getApplicationSystemTypesFromModel(p_oModel, aObjectsToClassify);   
}

/*---------------------------------------------------------------------------------------
    Check for each classified questionnaire template which of the found possible
    implicit placings of dependent sections and questions are really relevant 
 ---------------------------------------------------------------------------------------*/
function determineImplicitVisibilityDependencies() {
    var aQTs = convertHashSetToJSArray(g_questionnaireTemplateSet);
    for (var i=0; i<aQTs.length; i++) {
        var aAcceptedSectionPlacings = filterImplicitSectionPlacings(aQTs[i]);
        g_realImplicitSections.put(aQTs[i], aAcceptedSectionPlacings);   
        var aAcceptedQuestionPlacings = filterImplicitQuestionPlacings(aQTs[i]);
        g_realImplicitQuestions.put(aQTs[i], aAcceptedQuestionPlacings);
    }   
}

/*---------------------------------------------------------------------------------------
    Filters all possible implicit section placings that were found by previous
    classification algorithm. It accepts only those which:
    - refer to a section which is not already part of the explicitly positioned sections
    - have no activating context question
    - have at least one activating context question which is either one of the explicit 
      positioned questions or is one of the other implicit positioned questions of
      this questionnaire template
 ---------------------------------------------------------------------------------------*/ 
function filterImplicitSectionPlacings(p_oQuestionnaireTemplateObjDef) {
	
    var aAcceptedSectionPlacings = new Array();
    
    var oAcceptedContextQuestionsSet = new java.util.HashSet();
    var oExplicitSectionsForQTSet = new java.util.HashSet();
    
    var aExplicitQuestionsForQT = g_explicitPositionQuestions.get(p_oQuestionnaireTemplateObjDef);
    if (aExplicitQuestionsForQT != null) {
        oAcceptedContextQuestionsSet = convertJSArrayToHashSet(aExplicitQuestionsForQT);
    }
    var aExplicitSectionsForQT = g_explicitPositionSections.get(p_oQuestionnaireTemplateObjDef);
    if (aExplicitSectionsForQT != null) {
        oExplicitSectionsForQTSet = convertJSArrayToHashSet(aExplicitSectionsForQT);
    }
    
    var aPlacingsToFilter = new Array();
    var aPlacings;
    aPlacings = g_possibleImplicitTopLevelSections.get(p_oQuestionnaireTemplateObjDef);
    if (aPlacings!= null) {aPlacingsToFilter = aPlacingsToFilter.concat(aPlacings);}
    aPlacings = g_possibleImplicitSections.get(p_oQuestionnaireTemplateObjDef);
    if (aPlacings!= null) {aPlacingsToFilter = aPlacingsToFilter.concat(aPlacings);}
    
    var bContinueFiltering = true;
    while (bContinueFiltering && aPlacingsToFilter.length > 0) {
        bContinueFiltering = false;
        var aPlacingsStillToCheck = new Array();
        for (var i=0; i<aPlacingsToFilter.length; i++) {
            //remove all entries which refer to sections which are already explicitly placed within this questionnaire template
            if (oExplicitSectionsForQTSet.contains(aPlacingsToFilter[i].oObjDefToPlace)) {
                bContinueFiltering = true;
            }
            //accept all entries which have no activation context questions
            else if (aPlacingsToFilter[i].aContextQuestions == null || aPlacingsToFilter[i].aContextQuestions.length == 0) {
                aAcceptedSectionPlacings.push(aPlacingsToFilter[i]);
                bContinueFiltering = true;
            }
            else {
                //check the activation context questions
                var bAtLeastOneContextQuestionAccepted = false;
                for (var j=0; j<aPlacingsToFilter[i].aContextQuestions.length; j++) {
                    if (oAcceptedContextQuestionsSet.contains(aPlacingsToFilter[i].aContextQuestions[j])) {
                        bAtLeastOneContextQuestionAccepted = true;
                        break;
                    }   
                }
                //accept all entries where at least one context question is amongst the meanwhile accepted context questions
                if (bAtLeastOneContextQuestionAccepted) {
                    aAcceptedSectionPlacings.push(aPlacingsToFilter[i]);
                    bContinueFiltering = true;
                    addAcceptedQuestionsRecursive(aPlacingsToFilter[i].oObjDefToPlace, oAcceptedContextQuestionsSet);
                }
                //keep all entries where this is not the case ready for the next filter iteration
                else {
                    aPlacingsStillToCheck.push(aPlacingsToFilter[i]);
                }
            }
        }
        aPlacingsToFilter = new Array();
        aPlacingsToFilter = aPlacingsToFilter.concat(aPlacingsStillToCheck);
    }
    
    //make sure that all implicitly placed sections will be part of the export
    for (var k=0; k<aAcceptedSectionPlacings.length; k++) {
        g_sectionSet.add(aAcceptedSectionPlacings[k].oObjDefToPlace);
    }
    
    return aAcceptedSectionPlacings;
}

function addAcceptedQuestionsRecursive(p_oSectionDef, oAcceptedContextQuestionsSet) {  
    var aQuestions = g_section2questionsHashMap.get(p_oSectionDef);
    if (aQuestions != null) {
        for (var i=0; i<aQuestions.length; i++) {
            oAcceptedContextQuestionsSet.add(aQuestions[i]);
        }
    } 
    var aSubSections = g_section2subsectionsHashMap.get(p_oSectionDef);
    if (aSubSections != null) {
        for (var j=0; j<aSubSections.length; j++) {
            addQuestionsRecursive(aSubSections[j], oAcceptedContextQuestionsSet);
        }
    }
}

/*---------------------------------------------------------------------------------------
    Filters all possible implicit question placings that were found by previous
    classification algorithm. It accepts only those which:
    - refer to a question which is not already part of the explicitly positioned sections
    - have no activating context question
    - have at least one activating context question which is either one of the explicit 
      positioned questions or is one of the other implicit positioned questions of
      this questionnaire template
 ---------------------------------------------------------------------------------------*/ 
function filterImplicitQuestionPlacings(p_oQuestionnaireTemplateObjDef) {
	
    var aAcceptedQuestionPlacings = new Array();
    
    var oAcceptedContextQuestionsSet = new java.util.HashSet();
    var oExplicitQuestionsForQTSet = new java.util.HashSet();
    
    var aExplicitQuestionsForQT = g_explicitPositionQuestions.get(p_oQuestionnaireTemplateObjDef);
    if (aExplicitQuestionsForQT != null) {
        oAcceptedContextQuestionsSet = convertJSArrayToHashSet(aExplicitQuestionsForQT);
        oExplicitQuestionsForQTSet = convertJSArrayToHashSet(aExplicitQuestionsForQT);
    }
    
    var aPlacingsToFilter = new Array();
    var aPlacings;
    aPlacings = g_possibleImplicitQuestions.get(p_oQuestionnaireTemplateObjDef);
    if (aPlacings!= null) {aPlacingsToFilter = aPlacingsToFilter.concat(aPlacings);}
    
    var bContinueFiltering = true;
    while (bContinueFiltering && aPlacingsToFilter.length > 0) {
        bContinueFiltering = false;
        var aPlacingsStillToCheck = new Array();
        for (var i=0; i<aPlacingsToFilter.length; i++) {
            //remove all entries which refer to questions which are already explicitly placed within this questionnaire template
            if (oExplicitQuestionsForQTSet.contains(aPlacingsToFilter[i].oObjDefToPlace)) {
                bContinueFiltering = true;
            }
            //accept all entries which have no activation context questions
            else if (aPlacingsToFilter[i].aContextQuestions == null || aPlacingsToFilter[i].aContextQuestions.length == 0) {
                aAcceptedQuestionPlacings.push(aPlacingsToFilter[i]);
                bContinueFiltering = true;
                oAcceptedContextQuestionsSet.add(aPlacingsToFilter[i].oObjDefToPlace);
            }
            else {
                //check the activation context questions
                var bAtLeastOneContextQuestionAccepted = false;
                for (var j=0; j<aPlacingsToFilter[i].aContextQuestions.length; j++) {
                    if (oAcceptedContextQuestionsSet.contains(aPlacingsToFilter[i].aContextQuestions[j])) {
                        bAtLeastOneContextQuestionAccepted = true;
                        break;
                    }   
                }
                //accept all entries where at least one context question is amongst the meanwhile accepted context questions
                if (bAtLeastOneContextQuestionAccepted) {
                    aAcceptedQuestionPlacings.push(aPlacingsToFilter[i]);
                    bContinueFiltering = true;
                    oAcceptedContextQuestionsSet.add(aPlacingsToFilter[i].oObjDefToPlace);
                }
                //keep all entries where this is not the case ready for the next filter iteration
                else {
                    aPlacingsStillToCheck.push(aPlacingsToFilter[i]);
                }
            }
        }
        aPlacingsToFilter = new Array();
        aPlacingsToFilter = aPlacingsToFilter.concat(aPlacingsStillToCheck);
    }
    
    //make sure that all implicitly placed questions will be part of the export
    for (var k=0; k<aAcceptedQuestionPlacings.length; k++) {
        g_questionSet.add(aAcceptedQuestionPlacings[k].oObjDefToPlace);
    }
    
    return aAcceptedQuestionPlacings;
}


/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ~~~~~~~ Restore HashSets and HashMaps from report context if possible ~~~~~~~
                       ~~~~~~~ Audit management ~~~~~~~
  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/

function restoreAuditManagementFromReportContext() {

	if ( Context.getProperty("g_auditManagementHashSetsFilled") == null 
         || !Context.getProperty("g_auditManagementHashSetsFilled").equals("true") ) {
         return false;
    }
    
	g_deactivatedLinkSet = restoreStringHashSetFromProperty("g_deactivatedLinkSet", ArisData.getActiveDatabase());
    g_userGroupSet = restoreHashSetFromProperty("g_userGroupSet", ArisData.getActiveDatabase());
    g_userSet = restoreHashSetFromProperty("g_userSet", ArisData.getActiveDatabase());
	
    g_auditTemplateSet = restoreHashSetFromProperty("g_auditTemplateSet", ArisData.getActiveDatabase());
    g_auditStepTemplateSet = restoreHashSetFromProperty("g_auditStepTemplateSet", ArisData.getActiveDatabase());

    
    g_auditTemplate2AuditStepTemplateHashMap = restoreHashMapFromProperty("g_auditTemplate2AuditStepTemplateHashMap", ArisData.getActiveDatabase());
    g_auditStepTemplate2AuditTemplateHashMap = restoreHashMapFromProperty("g_auditStepTemplate2AuditTemplateHashMap", ArisData.getActiveDatabase());
    g_auditStepTemplate2subAuditStepTemplateHashMap = restoreHashMapFromProperty("g_auditStepTemplate2subAuditStepTemplateHashMap", ArisData.getActiveDatabase());
	g_auditStepTemplate2topAuditStepTemplateHashMap = restoreHashMapFromProperty("g_auditStepTemplate2topAuditStepTemplateHashMap", ArisData.getActiveDatabase());
    g_auditTemplate2auditOwnerHashMap = restoreHashMapFromProperty("g_auditTemplate2auditOwnerHashMap", ArisData.getActiveDatabase());
    g_auditTemplate2auditReviewerHashMap = restoreHashMapFromProperty("g_auditTemplate2auditReviewerHashMap", ArisData.getActiveDatabase());
    g_auditTemplate2auditAuditorHashMap = restoreHashMapFromProperty("g_auditTemplate2auditAuditorHashMap", ArisData.getActiveDatabase());
    g_auditStepTemplate2auditStepOwnerHashMap = restoreHashMapFromProperty("g_auditStepTemplate2auditStepOwnerHashMap", ArisData.getActiveDatabase());
    
	g_auditTemplate2regulationHashMap = restoreHashMapFromProperty("g_auditTemplate2regulationHashMap", ArisData.getActiveDatabase());
    g_auditTemplate2orgUnitHashMap = restoreHashMapFromProperty("g_auditTemplate2orgUnitHashMap", ArisData.getActiveDatabase());
    g_auditTemplate2processHashMap = restoreHashMapFromProperty("g_auditTemplate2processHashMap", ArisData.getActiveDatabase());
    g_auditTemplate2appSysTypeHashMap = restoreHashMapFromProperty("g_auditTemplate2appSysTypeHashMap", ArisData.getActiveDatabase());
    g_auditTemplate2riskCategoryHashMap = restoreHashMapFromProperty("g_auditTemplate2riskCategoryHashMap", ArisData.getActiveDatabase());
	
    g_auditStepTemplate2regulationHashMap = restoreHashMapFromProperty("g_auditStepTemplate2regulationHashMap", ArisData.getActiveDatabase());
    g_auditStepTemplate2orgUnitHashMap = restoreHashMapFromProperty("g_auditStepTemplate2orgUnitHashMap", ArisData.getActiveDatabase());
    g_auditStepTemplate2processHashMap = restoreHashMapFromProperty("g_auditStepTemplate2processHashMap", ArisData.getActiveDatabase());
    g_auditStepTemplate2appSysTypeHashMap = restoreHashMapFromProperty("g_auditStepTemplate2appSysTypeHashMap", ArisData.getActiveDatabase());
    g_auditStepTemplate2riskCategoryHashMap = restoreHashMapFromProperty("g_auditStepTemplate2riskCategoryHashMap", ArisData.getActiveDatabase());
    
    return true;
}

function storeAuditManagementToReportContext() {

	if ( !g_auditManagementHashSetsFilled ) {return;}

	Context.setProperty("g_auditManagementHashSetsFilled", "true");
	
	storeStringHashSetAsProperty("g_deactivatedLinkSet", g_deactivatedLinkSet);
    storeHashSetAsProperty("g_userGroupSet", g_userGroupSet);
    storeHashSetAsProperty("g_userSet", g_userSet);
	
	storeHashSetAsProperty("g_auditTemplateSet", g_auditTemplateSet);
	storeHashSetAsProperty("g_auditStepTemplateSet", g_auditStepTemplateSet);
	
	storeHashMapAsProperty("g_auditTemplate2AuditStepTemplateHashMap", g_auditTemplate2AuditStepTemplateHashMap);
	storeHashMapAsProperty("g_auditStepTemplate2AuditTemplateHashMap", g_auditStepTemplate2AuditTemplateHashMap);
	storeHashMapAsProperty("g_auditStepTemplate2subAuditStepTemplateHashMap", g_auditStepTemplate2subAuditStepTemplateHashMap);
	storeHashMapAsProperty("g_auditStepTemplate2topAuditStepTemplateHashMap", g_auditStepTemplate2topAuditStepTemplateHashMap);
	storeHashMapAsProperty("g_auditTemplate2auditOwnerHashMap", g_auditTemplate2auditOwnerHashMap);
	storeHashMapAsProperty("g_auditTemplate2auditReviewerHashMap", g_auditTemplate2auditReviewerHashMap);
	storeHashMapAsProperty("g_auditTemplate2auditAuditorHashMap", g_auditTemplate2auditAuditorHashMap);
	storeHashMapAsProperty("g_auditStepTemplate2auditStepOwnerHashMap", g_auditStepTemplate2auditStepOwnerHashMap);
	
	storeHashMapAsProperty("g_auditTemplate2regulationHashMap", g_auditTemplate2regulationHashMap);
	storeHashMapAsProperty("g_auditTemplate2orgUnitHashMap", g_auditTemplate2orgUnitHashMap);
	storeHashMapAsProperty("g_auditTemplate2processHashMap", g_auditTemplate2processHashMap);
	storeHashMapAsProperty("g_auditTemplate2appSysTypeHashMap", g_auditTemplate2appSysTypeHashMap);
	storeHashMapAsProperty("g_auditTemplate2riskCategoryHashMap", g_auditTemplate2riskCategoryHashMap);
	
	storeHashMapAsProperty("g_auditStepTemplate2regulationHashMap", g_auditStepTemplate2regulationHashMap);
	storeHashMapAsProperty("g_auditStepTemplate2orgUnitHashMap", g_auditStepTemplate2orgUnitHashMap);
	storeHashMapAsProperty("g_auditStepTemplate2processHashMap", g_auditStepTemplate2processHashMap);
	storeHashMapAsProperty("g_auditStepTemplate2appSysTypeHashMap", g_auditStepTemplate2appSysTypeHashMap);
	storeHashMapAsProperty("g_auditStepTemplate2riskCategoryHashMap", g_auditStepTemplate2riskCategoryHashMap);
}
  
/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    ~~~~~~~ Determine start objects ~~~~~~~
                        ~~~~~~~ Audit management ~~~~~~~
  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
/*---------------------------------------------------------------------------------------
    Determine the start objects relevant for basic ARCM audit management depending on 
    the selected context: 
    - Group: 
    - Model: Project schedule, Organisational chart 
    - ObjDef: Audit template, Person type (User group)
	Processes the found audit templates only if the global flag "g_bAuditManagementExport" 
	is true (see determineMode() in arcm-common.js).
    Fills the corresponding global HashSets.
    Returns an array of all found relevant objects.
 ---------------------------------------------------------------------------------------*/ 
function getAuditManagementSelectionByContext() {
    
    Context.writeStatus("Searching audit management data");
    //start object search only if the object GUIDs are not already passed as Context properties
    if (!restoreAuditManagementFromReportContext()) {
        buildAllowedModels();        
		// context "Group"
		var oSelectedGroups = ArisData.getSelectedGroups();
        if(oSelectedGroups.length == 0){
            var databases = ArisData.getSelectedDatabases();
            if (databases.length != 0) {
                oSelectedGroups = [databases[0].RootGroup()];
            }
        }
		if (oSelectedGroups.length > 0) {
			for (var i = 0; i < oSelectedGroups.length; i++ ) {
				getAuditManagementRelevantObjects(oSelectedGroups[i]);
			}
		}
		
		// context "Model"    
		var oSelectedModels = ArisData.getSelectedModels();
		if (oSelectedModels.length > 0) {
			for (var i = 0; i < oSelectedModels.length; i++ ) {
				
				//model must be marked as export relevant
				if (isExportRelevant(oSelectedModels[i])) {
					classifyObjectsByAuditManagementModel(oSelectedModels[i]);
				}
			}
		}
		
		// context "ObjDef"    
		var oSelectedObjDefs = ArisData.getSelectedObjDefs();
        if (oSelectedObjDefs.length > 0) {
            for (var i = 0; i < oSelectedObjDefs.length; i++ ) {       
                if (oSelectedObjDefs[i].TypeNum() == Constants.OT_FUNC_INST && oSelectedObjDefs[i].getDefaultSymbolNum() == Constants.ST_PROJECT) {        
                    processAuditManagementAuditTemplate(oSelectedObjDefs[i]);
                }
            }
        }
		
		//after global HashSets are filled:
		//Find all users to all found user groups   
		classifyUsers();
        removeEntriesFromUrgUnitSetIfinTesterSet();
		
		g_auditManagementHashSetsFilled = true;
		storeAuditManagementToReportContext();
	}
}
  
/*---------------------------------------------------------------------------------------
    Determine the relevant objects in the given group
---------------------------------------------------------------------------------------*/
function getAuditManagementRelevantObjects(p_oGroup) {

	// Get export audit template objects on group
    var oAuditTemplates = p_oGroup.ObjDefListFilter(Constants.OT_FUNC_INST);
    for (var i = 0; i < oAuditTemplates.length; i++ ) {
		if (oAuditTemplates[i].getDefaultSymbolNum() == Constants.ST_PROJECT) {
            processAuditManagementAuditTemplate(oAuditTemplates[i]);
        }
    }
    
    // Get objects in export relevant models
    var oModels = p_oGroup.ModelList();
    for (var i = 0; i < oModels.length; i++ ) {

        //model must be marked as export relevant        
        if (isExportRelevant(oModels[i])) {
            classifyObjectsByAuditManagementModel(oModels[i]);
        }
    }
    
    //Search by recursion in child groups
    var oChildGroups = p_oGroup.Childs();
    for (var i = 0; i < oChildGroups.length; i++ ) {
        getAuditManagementRelevantObjects(oChildGroups[i]);
    }

}

/*---------------------------------------------------------------------------------------
    Classify objects of a given model according to audit management modeling
 ---------------------------------------------------------------------------------------*/
function classifyObjectsByAuditManagementModel(p_oModel) {

	var aObjectsToClassify = p_oModel.ObjDefList();
    
    //project schedule model
    if (p_oModel.TypeNum() == Constants.MT_PROJECT_SCHEDULE) {
        for (var i=0; i<aObjectsToClassify.length; i++) {
            if (aObjectsToClassify[i].TypeNum() == Constants.OT_FUNC_INST && aObjectsToClassify[i].getDefaultSymbolNum() == Constants.ST_PROJECT ) {
                processAuditManagementAuditTemplate(aObjectsToClassify[i]);
            }
        }
    }
    
    //organisational charts
    else if (p_oModel.TypeNum() == Constants.MT_ORG_CHRT) {
        for (var i=0; i<aObjectsToClassify.length; i++) {
            //person types -> user groups
            if (aObjectsToClassify[i].TypeNum() == Constants.OT_PERS_TYPE) {
                //only "real" usergroups, not their role objects
                if (getRoleObject(aObjectsToClassify[i]) == null) {
                    continue;
                }
                g_userGroupSet.add(aObjectsToClassify[i]);
            }
        }    
    }

	getRiskCategoriesFromModel(p_oModel, aObjectsToClassify);
    getOrgUnitsFromModel(p_oModel, aObjectsToClassify);
	getRegulationsFromModel(p_oModel, aObjectsToClassify);    
    getProcessesFromModel(p_oModel, aObjectsToClassify);    
    getApplicationSystemTypesFromModel(p_oModel, aObjectsToClassify);
}  

/*---------------------------------------------------------------------------------------
    Determines the objects that are linked to the given audit template:
    - Audit step template
	- Audit owner
	- Audit reviewer
	- Audit auditor
 ---------------------------------------------------------------------------------------*/
function processAuditManagementAuditTemplate(p_oAuditTemplateDef) {

	//if the global flag "g_bAuditManagementExport" is not set then don't process anything
	if (!g_bAuditManagementExport) {
		return;
	}

	//audit template must be marked as export relevant
    if (isExportRelevant(p_oAuditTemplateDef)) {
    
		//if audit template was handled before then skip it, else process it further
		if ( !g_auditTemplateSet.add(p_oAuditTemplateDef) ) {
			return;    
		}

        // Audit Template --> Audit Step Template
        var aAllAssignedAuditStepTemplateDefs = getConnectedObjects(p_oAuditTemplateDef, Constants.EDGES_ASSIGN, Constants.CT_IS_PRCS_ORNT_SUPER, Constants.OT_FUNC_INST, Constants.ST_FUNC_INST);  
        //since all step templates have implicit connections to the audit template we must sort out the real top level step templates
		//-> all step templates which are not connected to any superior step template are top level
		var aTopLevelAuditStepTemplateDefs = new determineTopLevelAuditStepTemplates(aAllAssignedAuditStepTemplateDefs);
		
		// Audit Template --> PersType        
		var aPersTypeDefs = getConnectedObjects(p_oAuditTemplateDef, Constants.EDGES_IN, Constants.CT_EXEC_2, Constants.OT_PERS_TYPE);
		//differentiate the roles
		var aAuditOwners = new Array();
		var aAuditReviewers = new Array();
		var aAuditAuditors = new Array();
		for (var j=0; j<aPersTypeDefs.length; j++) {
			if (isPersonTypeAuditOwner(aPersTypeDefs[j])) {aAuditOwners.push(aPersTypeDefs[j]);}
			if (isPersonTypeAuditReviewer(aPersTypeDefs[j])) {aAuditReviewers.push(aPersTypeDefs[j]);}
			if (isPersonTypeAuditAuditor(aPersTypeDefs[j])) {aAuditAuditors.push(aPersTypeDefs[j]);}
		}
		
		// Audit Template --> Scope
		// process
		var aProcessScope = getConnectedObjects(p_oAuditTemplateDef, Constants.EDGES_IN, Constants.CT_IS_WITHIN_THE_SCOPE_OF, Constants.OT_FUNC);
		for (var i=0; i<aProcessScope.length; i++) {
            getFurtherObjectsByFunc(aProcessScope[i]);
		}		
		//orgunit
		var aOrgUnitScope = getConnectedObjects(p_oAuditTemplateDef, Constants.EDGES_IN, Constants.CT_IS_WITHIN_THE_SCOPE_OF, Constants.OT_ORG_UNIT);
		for (var i=0; i<aOrgUnitScope.length; i++) {	
            getFurtherOrgUnitHierarchyObjectsByOrgUnitDef(aOrgUnitScope[i]);
		}
		//regulation
		var aRegulationScope = getConnectedObjects(p_oAuditTemplateDef, Constants.EDGES_IN, Constants.CT_IS_WITHIN_THE_SCOPE_OF, Constants.OT_TECH_TRM);
		for (var i=0; i<aRegulationScope.length; i++) {
            getFurtherObjectsByRegulation(aRegulationScope[i]);
		}
		//appsystype
		var aAppSysTypeScope = getConnectedObjects(p_oAuditTemplateDef, Constants.EDGES_IN, Constants.CT_IS_WITHIN_THE_SCOPE_OF, Constants.OT_APPL_SYS_TYPE);
		for (var i=0; i<aAppSysTypeScope.length; i++) {
            getFurtherObjectsByAppSysType(aAppSysTypeScope[i]);
		}	
		//risk category
		var aRiskCategoryScope = getConnectedObjects(p_oAuditTemplateDef, Constants.EDGES_IN, Constants.CT_IS_WITHIN_THE_SCOPE_OF, Constants.OT_RISK_CATEGORY);
		for (var i=0; i<aRiskCategoryScope.length; i++) {
    		getFurtherObjectsByRiskCategory(aRiskCategoryScope[i]);
		}

			
        //update global HashMap for Audit Step Templates
        g_auditTemplate2AuditStepTemplateHashMap.put(p_oAuditTemplateDef, aTopLevelAuditStepTemplateDefs);       
        for (var j=0; j<aTopLevelAuditStepTemplateDefs.length; j++) {
            var oAuditStepTemplateDef = aTopLevelAuditStepTemplateDefs[j];

            var aAuditTemplates = g_auditStepTemplate2AuditTemplateHashMap.get(oAuditStepTemplateDef);
            if (aAuditTemplates == null) {aAuditTemplates = new Array();}
            aAuditTemplates.push(p_oAuditTemplateDef);
            g_auditStepTemplate2AuditTemplateHashMap.put(oAuditStepTemplateDef, aAuditTemplates);
            
            // Audit step template --> ...
            // add current audit step template to global set  
            if (g_auditStepTemplateSet.add(oAuditStepTemplateDef)) {               
                getFurtherObjectsByAuditStepTemplate(oAuditStepTemplateDef);
            }
        }
		
		//update global HashMap for Audit Owners
		g_auditTemplate2auditOwnerHashMap.put(p_oAuditTemplateDef, aAuditOwners); 
		for (var i=0; i<aAuditOwners.length ;i++) {g_userGroupSet.add(aAuditOwners[i]);} // add group to global set
		
		//update global HashMap for Audit Reviewers
		g_auditTemplate2auditReviewerHashMap.put(p_oAuditTemplateDef, aAuditReviewers); 
		for (var i=0; i<aAuditReviewers.length ;i++) {g_userGroupSet.add(aAuditReviewers[i]);} // add group to global set
		
		//update global HashMap for Audit Auditors
		g_auditTemplate2auditAuditorHashMap.put(p_oAuditTemplateDef, aAuditAuditors); 
		for (var i=0; i<aAuditAuditors.length ;i++) {g_userGroupSet.add(aAuditAuditors[i]);} // add group to global set
    
		//update global HashMaps for Audit template scope assignment
		g_auditTemplate2processHashMap.put(p_oAuditTemplateDef, aProcessScope); 
		g_auditTemplate2orgUnitHashMap.put(p_oAuditTemplateDef, aOrgUnitScope);
		g_auditTemplate2regulationHashMap.put(p_oAuditTemplateDef, aRegulationScope);
		g_auditTemplate2appSysTypeHashMap.put(p_oAuditTemplateDef, aAppSysTypeScope);
		g_auditTemplate2riskCategoryHashMap.put(p_oAuditTemplateDef, aRiskCategoryScope);
    }
}

/*---------------------------------------------------------------------------------------
    Determines all audit step templates which are not connected to a superior one.
 ---------------------------------------------------------------------------------------*/
function determineTopLevelAuditStepTemplates(aAuditStepTemplates) {

	var aTopLevelAuditStepTemplateDefs = new Array();

	for (var i=0; i<aAuditStepTemplates.length; i++) {
		var oAuditStepTemplateDef = aAuditStepTemplates[i];
		var aSuperiorAssignedAuditStepTemplateDefs = getConnectedObjects(oAuditStepTemplateDef, Constants.EDGES_OUT, Constants.CT_BELONGS_TO_10, Constants.OT_FUNC_INST, Constants.ST_FUNC_INST);
		if (aSuperiorAssignedAuditStepTemplateDefs.length == 0) {
			aTopLevelAuditStepTemplateDefs.push(oAuditStepTemplateDef);
		}
	}
	
	return aTopLevelAuditStepTemplateDefs;
}

/*---------------------------------------------------------------------------------------
    Determines the objects that are linked to the given audit step template:
    - Audit step template
	- Audit step owner
 ---------------------------------------------------------------------------------------*/
function getFurtherObjectsByAuditStepTemplate(p_oAuditStepTemplateDef) {

	// Audit Step Template --> Audit Step Template
    var aSubordinatedAuditStepTemplateDefs = getConnectedObjects(p_oAuditStepTemplateDef, Constants.EDGES_IN, Constants.CT_BELONGS_TO_10, Constants.OT_FUNC_INST, Constants.ST_FUNC_INST);  
    
    // Audit Template --> PersType        
	var aPersTypeDefs = getConnectedObjects(p_oAuditStepTemplateDef, Constants.EDGES_IN, Constants.CT_EXEC_2, Constants.OT_PERS_TYPE);
	var aAuditStepOwners = new Array();
	for (var j=0; j<aPersTypeDefs.length; j++) {
		if (isPersonTypeAuditStepOwner(aPersTypeDefs[j])) {aAuditStepOwners.push(aPersTypeDefs[j]);}
	}
	
	// Audit Step Template --> Scope
	// process
	var aProcessScope = getConnectedObjects(p_oAuditStepTemplateDef, Constants.EDGES_IN, Constants.CT_IS_WITHIN_THE_SCOPE_OF, Constants.OT_FUNC);
	for (var i=0; i<aProcessScope.length; i++) {
        getFurtherObjectsByFunc(aProcessScope[i]);
	}		
	//orgunit
	var aOrgUnitScope = getConnectedObjects(p_oAuditStepTemplateDef, Constants.EDGES_IN, Constants.CT_IS_WITHIN_THE_SCOPE_OF, Constants.OT_ORG_UNIT);
	for (var i=0; i<aOrgUnitScope.length; i++) {	
        getFurtherOrgUnitHierarchyObjectsByOrgUnitDef(aOrgUnitScope[i]);
	}
	//regulation
	var aRegulationScope = getConnectedObjects(p_oAuditStepTemplateDef, Constants.EDGES_IN, Constants.CT_IS_WITHIN_THE_SCOPE_OF, Constants.OT_TECH_TRM);
	for (var i=0; i<aRegulationScope.length; i++) {
        getFurtherObjectsByRegulation(aRegulationScope[i]);
	}
	//appsystype
	var aAppSysTypeScope = getConnectedObjects(p_oAuditStepTemplateDef, Constants.EDGES_IN, Constants.CT_IS_WITHIN_THE_SCOPE_OF, Constants.OT_APPL_SYS_TYPE);
	for (var i=0; i<aAppSysTypeScope.length; i++) {
        getFurtherObjectsByAppSysType(aAppSysTypeScope[i]);
	}	
	//risk category
	var aRiskCategoryScope = getConnectedObjects(p_oAuditStepTemplateDef, Constants.EDGES_IN, Constants.CT_IS_WITHIN_THE_SCOPE_OF, Constants.OT_RISK_CATEGORY);
	for (var i=0; i<aRiskCategoryScope.length; i++) {
        getFurtherObjectsByRiskCategory(aRiskCategoryScope[i]);
	}
	
	
	//update global HashMap for Audit Step Templates -> Sub Audit Step Templates
	g_auditStepTemplate2subAuditStepTemplateHashMap.put(p_oAuditStepTemplateDef, aSubordinatedAuditStepTemplateDefs);       
	
	for (var j=0; j<aSubordinatedAuditStepTemplateDefs.length; j++) {	
		var oSubAuditStepTemplateDef = aSubordinatedAuditStepTemplateDefs[j];
		
		//update global HashMap for Audit Step Templates -> Top Audit Step Templates
		var aTopSteps = g_auditStepTemplate2topAuditStepTemplateHashMap.get(oSubAuditStepTemplateDef);
		if (aTopSteps == null) {aTopSteps = new Array();}
		aTopSteps.push(p_oAuditStepTemplateDef);
		g_auditStepTemplate2topAuditStepTemplateHashMap.put(oSubAuditStepTemplateDef, aTopSteps);
		
		// Audit step template --> ...
        if (g_auditStepTemplateSet.add(oSubAuditStepTemplateDef)) {
            getFurtherObjectsByAuditStepTemplate(oSubAuditStepTemplateDef);
        }
	}
	
	//update global HashMap for Audit Owners
	g_auditStepTemplate2auditStepOwnerHashMap.put(p_oAuditStepTemplateDef, aAuditStepOwners); 
	for (var i=0; i<aAuditStepOwners.length ;i++) {g_userGroupSet.add(aAuditStepOwners[i]);} // add group to global set
	
	//update global HashMaps for Audit step template scope assignment
	g_auditStepTemplate2processHashMap.put(p_oAuditStepTemplateDef, aProcessScope); 
	g_auditStepTemplate2orgUnitHashMap.put(p_oAuditStepTemplateDef, aOrgUnitScope);
	g_auditStepTemplate2regulationHashMap.put(p_oAuditStepTemplateDef, aRegulationScope);
	g_auditStepTemplate2appSysTypeHashMap.put(p_oAuditStepTemplateDef, aAppSysTypeScope);
	g_auditStepTemplate2riskCategoryHashMap.put(p_oAuditStepTemplateDef, aRiskCategoryScope);
}

/*---------------------------------------------------------------------------------------
    Checks for objects of type "User group" if they are audit owners.
    For standard mapping: They are if they have one assigned role object which name starts 
    with "Audit owner".
 ---------------------------------------------------------------------------------------*/        
function isPersonTypeAuditOwner(p_userGroupDef) {
    return hasARCMRole(p_userGroupDef, "auditowner");
}

/*---------------------------------------------------------------------------------------
    Checks for objects of type "User group" if they are audit reviewers.
    For standard mapping: They are if they have one assigned role object which name starts 
    with "Audit reviewer".
 ---------------------------------------------------------------------------------------*/        
function isPersonTypeAuditReviewer(p_userGroupDef) {
    return hasARCMRole(p_userGroupDef, "auditreviewer");
}

/*---------------------------------------------------------------------------------------
    Checks for objects of type "User group" if they are audit managers.
    For standard mapping: They are if they have one assigned role object which name starts 
    with "Audit manager".
 ---------------------------------------------------------------------------------------*/        
function isPersonTypeAuditManager(p_userGroupDef) {
    return hasARCMRole(p_userGroupDef, "auditmanager");
}

/*---------------------------------------------------------------------------------------
    Checks for objects of type "User group" if they are audit step owners.
    For standard mapping: They are if they have one assigned role object which name starts 
    with "Audit step owner".
 ---------------------------------------------------------------------------------------*/        
function isPersonTypeAuditStepOwner(p_userGroupDef) {
    return hasARCMRole(p_userGroupDef, "auditstepowner");
}

/*---------------------------------------------------------------------------------------
    Checks for objects of type "User group" if they are audit auditor.
    For standard mapping: They are if they have one assigned role object which name starts 
    with "Audit auditor".
 ---------------------------------------------------------------------------------------*/        
function isPersonTypeAuditAuditor(p_userGroupDef) {
    return hasARCMRole(p_userGroupDef, "auditauditor");
}


/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ~~~~~~~ Restore HashSets and HashMaps from report context if possible ~~~~~~~
                      ~~~~~~~ Policy management ~~~~~~~
  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
/**
  * Restore all policy management HashSets and HashMaps
  */
function restorePolicyManagementFromReportContext() {
    
	if ( Context.getProperty("g_policyManagementHashSetsFilled") == null 
         || !Context.getProperty("g_policyManagementHashSetsFilled").equals("true") ) {
         return false;
    }
    
    g_deactivatedLinkSet = restoreStringHashSetFromProperty("g_deactivatedLinkSet", ArisData.getActiveDatabase());
    g_userGroupSet = restoreHashSetFromProperty("g_userGroupSet", ArisData.getActiveDatabase());
    g_userSet = restoreHashSetFromProperty("g_userSet", ArisData.getActiveDatabase());
    
    g_policyDefinitionSet = restoreHashSetFromProperty("g_policyDefinitionSet", ArisData.getActiveDatabase());

    g_policyDefinition2policyOwnerHashMap = restoreHashMapFromProperty("g_policyDefinition2policyOwnerHashMap", ArisData.getActiveDatabase());
    g_policyDefinition2policyApproverHashMap = restoreHashMapFromProperty("g_policyDefinition2policyApproverHashMap", ArisData.getActiveDatabase());
    g_policyDefinition2policyAddresseeHashMap = restoreHashMapFromProperty("g_policyDefinition2policyAddresseeHashMap", ArisData.getActiveDatabase());
    g_policyDefinition2policyReviewerHashMap = restoreHashMapFromProperty("g_policyDefinition2policyReviewerHashMap", ArisData.getActiveDatabase());
    g_policyDefinition2policyAuditorHashMap = restoreHashMapFromProperty("g_policyDefinition2policyAuditorHashMap", ArisData.getActiveDatabase());

    g_policyDefinition2riskHashMap = restoreHashMapFromProperty("g_policyDefinition2riskHashMap", ArisData.getActiveDatabase());
    g_policyDefinition2regulationHashMap = restoreHashMapFromProperty("g_policyDefinition2regulationHashMap", ArisData.getActiveDatabase());
    g_policyDefinition2orgUnitHashMap = restoreHashMapFromProperty("g_policyDefinition2orgUnitHashMap", ArisData.getActiveDatabase());
    g_policyDefinition2processHashMap = restoreHashMapFromProperty("g_policyDefinition2processHashMap", ArisData.getActiveDatabase());
    g_policyDefinition2appSysTypeHashMap = restoreHashMapFromProperty("g_policyDefinition2appSysTypeHashMap", ArisData.getActiveDatabase());
    
    return true;
}

function storePolicyManagementToReportContext() {

	if ( !g_policyManagementHashSetsFilled ) {return;}

	Context.setProperty("g_policyManagementHashSetsFilled", "true");
	
	storeStringHashSetAsProperty("g_deactivatedLinkSet", g_deactivatedLinkSet);
    storeHashSetAsProperty("g_userGroupSet", g_userGroupSet);
    storeHashSetAsProperty("g_userSet", g_userSet);
	
	storeHashSetAsProperty("g_policyDefinitionSet", g_policyDefinitionSet);
	
	storeHashMapAsProperty("g_policyDefinition2policyOwnerHashMap", g_policyDefinition2policyOwnerHashMap);
	storeHashMapAsProperty("g_policyDefinition2policyApproverHashMap", g_policyDefinition2policyApproverHashMap);
	storeHashMapAsProperty("g_policyDefinition2policyAddresseeHashMap", g_policyDefinition2policyAddresseeHashMap);
	storeHashMapAsProperty("g_policyDefinition2policyReviewerHashMap", g_policyDefinition2policyReviewerHashMap);
	storeHashMapAsProperty("g_policyDefinition2policyAuditorHashMap", g_policyDefinition2policyAuditorHashMap);
	
    storeHashMapAsProperty("g_policyDefinition2riskHashMap", g_policyDefinition2riskHashMap);
	storeHashMapAsProperty("g_policyDefinition2regulationHashMap", g_policyDefinition2regulationHashMap);
	storeHashMapAsProperty("g_policyDefinition2orgUnitHashMap", g_policyDefinition2orgUnitHashMap);
	storeHashMapAsProperty("g_policyDefinition2processHashMap", g_policyDefinition2processHashMap);
	storeHashMapAsProperty("g_policyDefinition2appSysTypeHashMap", g_policyDefinition2appSysTypeHashMap);
}

/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    ~~~~~~~ Determine start objects ~~~~~~~
                       ~~~~~~~ Policy management ~~~~~~~
  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
 /*---------------------------------------------------------------------------------------
	Determine the start objects relevant for basic ARCM policy management depending on 
    the selected context: 
    - Group: 
    - Model: Business controls diagram, EPC, EPC (column display), EPC (material flow),
			 EPC (row display), EPC (table display), EPC (horizontal table display),
			 Function allocation diagram, PCD (material flow), PCD, Value-added chain diagram
    - ObjDef: Policy
	Processes the found policies only if the global flag "g_bPolicyManagementExport" 
	is true (see determineMode() in arcm-common.js).
    Fills the corresponding global HashSets.
    Returns an array of all found relevant objects.
 ---------------------------------------------------------------------------------------*/
function getPolicyManagementSelectionByContext() {
    
    Context.writeStatus("Searching policy management data");
    //start object search only if the object GUIDs are not already passed as Context properties
    if (!restorePolicyManagementFromReportContext()) {        
        buildAllowedModels();
		// context "Group"
		var oSelectedGroups = ArisData.getSelectedGroups();
        if(oSelectedGroups.length == 0){
            var databases = ArisData.getSelectedDatabases();
            if (databases.length != 0) {
                oSelectedGroups = [databases[0].RootGroup()];
            }
        }
		if (oSelectedGroups.length > 0) {
			for (var i = 0; i < oSelectedGroups.length; i++ ) {
				getPolicyManagementRelevantObjects(oSelectedGroups[i]);
			}
		}
		
		// context "Model"    
		var oSelectedModels = ArisData.getSelectedModels();
		if (oSelectedModels.length > 0) {
			for (var i = 0; i < oSelectedModels.length; i++ ) {
				
				//model must be marked as export relevant
				if (isExportRelevant(oSelectedModels[i])) {
					classifyObjectsByPolicyManagementModel(oSelectedModels[i]);
				}
			}
		}
		
		// context "ObjDef"    
		var oSelectedObjDefs = ArisData.getSelectedObjDefs();
        if (oSelectedObjDefs.length > 0) {
            for (var i = 0; i < oSelectedObjDefs.length; i++ ) {       
                if (oSelectedObjDefs[i].TypeNum() == Constants.OT_POLICY) {        
                    processPolicyManagementPolicyDefinition(oSelectedObjDefs[i]);
                }
            }
        }
		
		//after global HashSets are filled:
		//Find all users to all found user groups   
		classifyUsers();
        removeEntriesFromUrgUnitSetIfinTesterSet();
		
		g_policyManagementHashSetsFilled = true;
		storePolicyManagementToReportContext();
	}
}
  
/*---------------------------------------------------------------------------------------
    Determine the relevant objects in the given group
---------------------------------------------------------------------------------------*/
function getPolicyManagementRelevantObjects(p_oGroup) {

	// Get export policy definition objects on group
    var oPolicyDefinitions = p_oGroup.ObjDefListFilter(Constants.OT_POLICY);
    for (var i = 0; i < oPolicyDefinitions.length; i++ ) {
        processPolicyManagementPolicyDefinition(oPolicyDefinitions[i]);
    }
    
    // Get objects in export relevant models
    var oModels = p_oGroup.ModelList();
    for (var i = 0; i < oModels.length; i++ ) {

        //model must be marked as export relevant        
        if (isExportRelevant(oModels[i])) {
            classifyObjectsByPolicyManagementModel(oModels[i]);
        }
    }
    
    //Search by recursion in child groups
    var oChildGroups = p_oGroup.Childs();
    for (var i = 0; i < oChildGroups.length; i++ ) {
        getPolicyManagementRelevantObjects(oChildGroups[i]);
    }

}

/*---------------------------------------------------------------------------------------
    Classify objects of a given model according to policy management modeling
 ---------------------------------------------------------------------------------------*/
function classifyObjectsByPolicyManagementModel(p_oModel) {

	var aObjectsToClassify = p_oModel.ObjDefList();
	
	//all models where policy occurences can be created
	if (p_oModel.TypeNum() == Constants.MT_BUSY_CONTR_DGM             //Business controls diagram
        || p_oModel.TypeNum() == Constants.MT_EEPC                    //EPC
        || p_oModel.TypeNum() == Constants.MT_EEPC_COLUMN             //EPC (column display)
        || p_oModel.TypeNum() == Constants.MT_EEPC_MAT                //EPC (material flow)
        || p_oModel.TypeNum() == Constants.MT_EEPC_ROW                //EPC (row display)
		|| p_oModel.TypeNum() == Constants.MT_EEPC_TAB                //EPC (table display)
		|| p_oModel.TypeNum() == Constants.MT_EEPC_TAB_HORIZONTAL     //EPC (horizontal table display)
		|| p_oModel.TypeNum() == Constants.MT_FUNC_ALLOC_DGM          //Function allocation diagram
        || p_oModel.TypeNum() == Constants.MT_PCD_MAT                 //PCD (material flow)
		|| p_oModel.TypeNum() == Constants.MT_PRCS_CHN_DGM            //PCD    
		|| p_oModel.TypeNum() == Constants.MT_VAL_ADD_CHN_DGM         //Value-added chain diagram
		) {
		for (var i=0; i<aObjectsToClassify.length; i++) {
            if (aObjectsToClassify[i].TypeNum() == Constants.OT_POLICY) {
                processPolicyManagementPolicyDefinition(aObjectsToClassify[i]);
            }
        }
	}
	
	getRegulationsFromModel(p_oModel, aObjectsToClassify);	
    getOrgUnitsFromModel(p_oModel, aObjectsToClassify);   
    getProcessesFromModel(p_oModel, aObjectsToClassify);    
    getApplicationSystemTypesFromModel(p_oModel, aObjectsToClassify);
}

/*---------------------------------------------------------------------------------------
    Determines the objects that are linked to the given policy definition:
    - Policy owner
    - Policy approver
    - Policy addressee
    - Policy reviewer
    - Policy auditor
    - Risk
    - Function
    - Technical term
    - Organisational unit
    - Application system type
 ---------------------------------------------------------------------------------------*/   
function processPolicyManagementPolicyDefinition(p_oPolicyDefinitionDef) {
    
	//if the global flag "g_bPolicyManagementExport" is not set then don't process anything
	if (!g_bPolicyManagementExport) {
		return;
	}
	
    //Policy definition must be marked as export relevant
    if (isExportRelevant(p_oPolicyDefinitionDef)) {
        
        //if policy definition was handled before then skip it, else process it further
        if ( !g_policyDefinitionSet.add(p_oPolicyDefinitionDef) ) {
            return;    
        }
        
        // Policy Definition --> PersType        
        var aPersTypeDefs = getConnectedObjects(p_oPolicyDefinitionDef, Constants.EDGES_IN, Constants.CT_IS_TECH_RESP_3, Constants.OT_PERS_TYPE);
        //differentiate the roles
        var aPolicyOwners = new Array();
        var aPolicyApprovers = new Array();
        var aPolicyAddressees = new Array();
        var aPolicyReviewer = new Array();
        var aPolicyAuditor = new Array();
        for (var j=0; j<aPersTypeDefs.length; j++) {
            if (isPersonTypePolicyOwner(aPersTypeDefs[j])) {aPolicyOwners.push(aPersTypeDefs[j]);}
            if (isPersonTypePolicyApprover(aPersTypeDefs[j])) {aPolicyApprovers.push(aPersTypeDefs[j]);}
            if (isPersonTypePolicyAddressee(aPersTypeDefs[j])) {aPolicyAddressees.push(aPersTypeDefs[j]);}
            if (isPersonTypePolicyReviewer(aPersTypeDefs[j])) {aPolicyReviewer.push(aPersTypeDefs[j]);}
            if (isPersonTypePolicyAuditor(aPersTypeDefs[j])) {aPolicyAuditor.push(aPersTypeDefs[j]);}
        }
        
        // Policy Definition --> Risk 	- only if g_bTestManagementExport or g_bRiskManagementExport is true
        var aRisks = getConnectedObjects(p_oPolicyDefinitionDef, Constants.EDGES_IN, Constants.CT_IS_REDU_BY, Constants.OT_RISK);
		if (g_bTestManagementExport || g_bRiskManagementExport) {
			for (var i=0; i<aRisks.length; i++) {
				if (g_riskSet.add(aRisks[i])) {
					if (g_bTestManagementExport) {
						getFurtherTestManagementObjectsByRisk(aRisks[i]);
					}
					if (g_bRiskManagementExport) {
						getFurtherRiskManagementObjectsByRisk(aRisks[i]);
					}
				}
			}
        }
        
        // Policy Definition --> Process
        var aProcesses = getConnectedObjects(p_oPolicyDefinitionDef, Constants.EDGES_OUT, Constants.CT_AFFECTS, Constants.OT_FUNC);
        for (var i=0; i<aProcesses.length; i++) {
            getFurtherObjectsByFunc(aProcesses[i]);
        }		
        //Policy Definition --> Orgunit
        var aOrgUnits = getConnectedObjects(p_oPolicyDefinitionDef, Constants.EDGES_OUT, Constants.CT_AFFECTS, Constants.OT_ORG_UNIT);
        for (var i=0; i<aOrgUnits.length; i++) {	
            getFurtherOrgUnitHierarchyObjectsByOrgUnitDef(aOrgUnits[i]);
        }
        //Policy Definition --> Regulation
        var aRegulations = getConnectedObjects(p_oPolicyDefinitionDef, Constants.EDGES_OUT, Constants.CT_AFFECTS, Constants.OT_TECH_TRM);
        for (var i=0; i<aRegulations.length; i++) {
            getFurtherObjectsByRegulation(aRegulations[i]);
        }
        //Policy Definition --> Appsystype
        var aAppSysTypes = getConnectedObjects(p_oPolicyDefinitionDef, Constants.EDGES_OUT, Constants.CT_AFFECTS, Constants.OT_APPL_SYS_TYPE);
        for (var i=0; i<aAppSysTypes.length; i++) {
            getFurtherObjectsByAppSysType(aAppSysTypes[i]);
        }
    
        
        //update global HashMap for Policy Owners
        g_policyDefinition2policyOwnerHashMap.put(p_oPolicyDefinitionDef, aPolicyOwners); 
        for (var i=0; i<aPolicyOwners.length ;i++) {g_userGroupSet.add(aPolicyOwners[i]);} // add group to global set
        
        //update global HashMap for Policy Approvers
        g_policyDefinition2policyApproverHashMap.put(p_oPolicyDefinitionDef, aPolicyApprovers); 
        for (var i=0; i<aPolicyApprovers.length ;i++) {g_userGroupSet.add(aPolicyApprovers[i]);} // add group to global set
        
        //update global HashMap for Policy Addressees
        g_policyDefinition2policyAddresseeHashMap.put(p_oPolicyDefinitionDef, aPolicyAddressees); 
        for (var i=0; i<aPolicyAddressees.length ;i++) {g_userGroupSet.add(aPolicyAddressees[i]);} // add group to global set
        
        //update global HashMap for Policy Reviewers
        g_policyDefinition2policyReviewerHashMap.put(p_oPolicyDefinitionDef, aPolicyReviewer); 
        for (var i=0; i<aPolicyReviewer.length ;i++) {g_userGroupSet.add(aPolicyReviewer[i]);} // add group to global set
        
        //update global HashMap for Policy Auditors
        g_policyDefinition2policyAuditorHashMap.put(p_oPolicyDefinitionDef, aPolicyAuditor); 
        for (var i=0; i<aPolicyAuditor.length ;i++) {g_userGroupSet.add(aPolicyAuditor[i]);} // add group to global set
    
        //update global HashMaps for Audit template scope assignment
        g_policyDefinition2riskHashMap.put(p_oPolicyDefinitionDef, aRisks); 
        g_policyDefinition2regulationHashMap.put(p_oPolicyDefinitionDef, aProcesses);
        g_policyDefinition2orgUnitHashMap.put(p_oPolicyDefinitionDef, aOrgUnits);
        g_policyDefinition2processHashMap.put(p_oPolicyDefinitionDef, aRegulations);
        g_policyDefinition2appSysTypeHashMap.put(p_oPolicyDefinitionDef, aAppSysTypes);
    }
}

/*---------------------------------------------------------------------------------------
    Checks for objects of type "User group" if they are policy owners.
    For standard mapping: They are if they have one assigned role object which name starts 
    with "Policy owner".
 ---------------------------------------------------------------------------------------*/        
function isPersonTypePolicyOwner(p_userGroupDef) {  
    return hasARCMRole(p_userGroupDef, "policyowner");
}

/*---------------------------------------------------------------------------------------
    Checks for objects of type "User group" if they are policy approvers.
    For standard mapping: They are if they have one assigned role object which name starts 
    with "Policy approver".
 ---------------------------------------------------------------------------------------*/        
function isPersonTypePolicyApprover(p_userGroupDef) {
    return hasARCMRole(p_userGroupDef, "policyapprover");
}

/*---------------------------------------------------------------------------------------
    Checks for objects of type "User group" if they are policy addressees.
    For standard mapping: They are if they have one assigned role object which name starts 
    with "Policy addressee".
 ---------------------------------------------------------------------------------------*/        
function isPersonTypePolicyAddressee(p_userGroupDef) {
    return hasARCMRole(p_userGroupDef, "policyaddressee");
}

/*---------------------------------------------------------------------------------------
    Checks for objects of type "User group" if they are policy reviewers.
    For standard mapping: They are if they have one assigned role object which name starts 
    with "Policy reviewer".
 ---------------------------------------------------------------------------------------*/        
function isPersonTypePolicyReviewer(p_userGroupDef) {
    return hasARCMRole(p_userGroupDef, "policyreviewer");
}

/*---------------------------------------------------------------------------------------
    Checks for objects of type "User group" if they are policy auditors.
    For standard mapping: They are if they have one assigned role object which name starts 
    with "Policy auditor".
 ---------------------------------------------------------------------------------------*/        
function isPersonTypePolicyAuditor(p_userGroupDef) {
    return hasARCMRole(p_userGroupDef, "policyauditor");
}


/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                             ~~~~~~~ GDPR ~~~~~~~
   ~~~~~~~ Work with the start objects classified by previous modules   ~~~~~~~
  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
/**
  * Restore all GDPR HashSets and HashMaps
  */
function restoreGDPRFromReportContext() {
    
    if ( Context.getProperty("g_gdprHashSetsFilled") == null 
         || !Context.getProperty("g_gdprHashSetsFilled").equals("true") ) {
         return false;
    }
    
	g_deactivatedLinkSet = restoreStringHashSetFromProperty("g_deactivatedLinkSet", ArisData.getActiveDatabase());
    g_userGroupSet = restoreHashSetFromProperty("g_userGroupSet", ArisData.getActiveDatabase());
    g_userSet = restoreHashSetFromProperty("g_userSet", ArisData.getActiveDatabase());
	
	g_processSet = restoreHashSetFromProperty("g_processSet", ArisData.getActiveDatabase());
  
    g_orgUnitSet = restoreHashSetFromProperty("g_orgUnitSet", ArisData.getActiveDatabase());
	g_appSysSet = restoreHashSetFromProperty("g_appSysSet", ArisData.getActiveDatabase());
    
    g_dataClusterSet = restoreHashSetFromProperty("g_dataClusterSet", ArisData.getActiveDatabase());
    
    return true;
}

/**
  * Store all GDPR HashSets and HashMaps
  */
function storeGDPRToReportContext() {
    
    if ( !g_gdprHashSetsFilled ) {return;}
    
    Context.setProperty("g_gdprHashSetsFilled", "true");
    
	storeStringHashSetAsProperty("g_deactivatedLinkSet", g_deactivatedLinkSet);
    storeHashSetAsProperty("g_userGroupSet", g_userGroupSet);
    storeHashSetAsProperty("g_userSet", g_userSet);
	
	storeHashSetAsProperty("g_processSet", g_processSet);
	
    storeHashSetAsProperty("g_orgUnitSet", g_orgUnitSet);
    storeHashSetAsProperty("g_appSysSet", g_appSysSet);
    
    storeHashSetAsProperty("g_dataClusterSet", g_dataClusterSet);
}
  
/*---------------------------------------------------------------------------------------
    Based on the relevant objects that were already classified for the given starting
	context by the other modules:
	Read the GDPR relevant data.
 ---------------------------------------------------------------------------------------*/   
function getGDPRSelectionByContext() {
    
    Context.writeStatus("Searching GDPR data");
    //start object search only if the object GUIDs are not already passed as Context properties
    if (!restoreGDPRFromReportContext()) {
        buildAllowedModels();        
        var aProcessArray = convertHashSetToJSArray(g_processSet);
		for (var i = 0; i < aProcessArray.length; i++ ) {
			classifyGDPRDataByProcess(aProcessArray[i]);
		}


    // if add, here is recursion call of folders required
        // context "Group"
        var oSelectedGroups = ArisData.getSelectedGroups();
        if(oSelectedGroups.length == 0){
            var databases = ArisData.getSelectedDatabases();
            if (databases.length != 0) {
                oSelectedGroups = [databases[0].RootGroup()];
            }
        }
        
        if (oSelectedGroups.length > 0) {
            for (var i = 0; i < oSelectedGroups.length; i++ ) {
                getGDPRRelevantObjects(oSelectedGroups[i]);
            }
        }
        
   		// context "Model"    
		var oSelectedModels = ArisData.getSelectedModels();
		if (oSelectedModels.length > 0) {
			for (var i = 0; i < oSelectedModels.length; i++ ) {				
				//model must be marked as export relevant
				if (isExportRelevant(oSelectedModels[i])) {
					classifyObjectsByGDPR(oSelectedModels[i]);
				}
			}
		}
                
        //after global HashSets are filled:
        //Find all users to all found user groups   
        classifyUsers();
        removeEntriesFromUrgUnitSetIfinTesterSet();

        g_gdprHashSetsFilled = true;
        storeGDPRToReportContext();
    }
}


function getGDPRRelevantObjects(p_oGroup) {		
		
    // Get objects in export relevant models
    var oModels = p_oGroup.ModelList();
    for (var i = 0; i < oModels.length; i++ ) {

        //model must be marked as export relevant        
        if (isExportRelevant(oModels[i])) {
            //aRelevantObjects = aRelevantObjects.concat(oModels[i].ObjDefList());
            classifyObjectsByGDPR(oModels[i]);
        }
    }
    
    //Search by recursion in child groups
    var oChildGroups = p_oGroup.Childs();
    for (var i = 0; i < oChildGroups.length; i++ ) {
        getGDPRRelevantObjects(oChildGroups[i]);
    }
}

function classifyObjectsByGDPR(p_oModel) {    
    var aObjectsToClassify = p_oModel.ObjDefList();
    getDataClusterFromModel(p_oModel, aObjectsToClassify);   			      
}

/*---------------------------------------------------------------------------------------
    Classifies all GDPR related hierarchies that are linked to one of the classified 
    process hierarchy items
 ---------------------------------------------------------------------------------------*/
function classifyGDPRDataByProcess(p_processDef) {    

    //if the process was not checked for GDPR data yet
    if (g_processGDPRSet.add(p_processDef)) { 
        
        //check if the process is relevant for GDPR (i.e. is an processing activity)
        if (!isGDPRRelevant(p_processDef)) { return; }    
		
        //linked orgunits - controller
        var a_connectedControllerOrgUnitDefs = getConnectedObjects(p_processDef, Constants.EDGES_IN, Constants.CT_IS_TECH_RESP_1, Constants.OT_ORG_UNIT);
        // add organisational units to global set
        for (var i=0; i<a_connectedControllerOrgUnitDefs.length ;i++) {
            getFurtherOrgUnitHierarchyObjectsByOrgUnitDef(a_connectedControllerOrgUnitDefs[i]);
        }     
        
        //linked orgunits - processor
        var a_connectedProcessorOrgUnitDefs = getConnectedObjects(p_processDef, Constants.EDGES_IN, Constants.CT_EXEC_1, Constants.OT_ORG_UNIT);
        // add organisational units to global set
        for (var i=0; i<a_connectedProcessorOrgUnitDefs.length ;i++) {
            getFurtherOrgUnitHierarchyObjectsByOrgUnitDef(a_connectedProcessorOrgUnitDefs[i]);
        }
        
        //linked datacluster
        //--> hierarchy parent recursion for datacluster
        var a_connectedDataClusterDefs = new Array();
        a_connectedDataClusterDefs = a_connectedDataClusterDefs.concat( getConnectedObjects(p_processDef, Constants.EDGES_OUT, Constants.CT_READ_1, Constants.OT_CLST) );
        a_connectedDataClusterDefs = a_connectedDataClusterDefs.concat( getConnectedObjects(p_processDef, Constants.EDGES_OUT, Constants.CT_HAS_OUT, Constants.OT_CLST) );
        // add data cluster units to global set
        for (var i=0; i<a_connectedDataClusterDefs.length ;i++) {
            if (g_dataClusterSet.add(a_connectedDataClusterDefs[i])) {
                getFurtherDataClusterHierarchyObjectsByDataClusterDef(a_connectedDataClusterDefs[i]);
            }   
        }
        
        //linked appsystypes
        var a_connectedAppSysTypeDefs = getConnectedObjects(p_processDef, Constants.EDGES_IN, Constants.CT_CAN_SUPP_1, Constants.OT_APPL_SYS_TYPE);
        // add appsys types to global set
        for (var i=0; i<a_connectedAppSysTypeDefs.length ;i++) {
            getFurtherObjectsByAppSysType(a_connectedAppSysTypeDefs[i]);
        }                
    }  
}

/*---------------------------------------------------------------------------------------
    Determines the objects that are linked to the data cluster:
    - Superior data cluster (recursively)
 ---------------------------------------------------------------------------------------*/
function getFurtherDataClusterHierarchyObjectsByDataClusterDef(p_oDataClusterDef) {
    //Data cluster -> Superior data cluster
    var aSuperiorDataClusters = getConnectedObjects(p_oDataClusterDef, Constants.EDGES_IN | Constants.EDGES_ASSIGN, Constants.CT_CONS_OF_2, Constants.OT_CLST);
    updateHierarchyStructureHashMaps(p_oDataClusterDef, aSuperiorDataClusters);
    
    for (var i=0; i<aSuperiorDataClusters.length; i++) {
        //if data cluster was not handled before
        if ( g_dataClusterSet.add(aSuperiorDataClusters[i]) ){
            getFurtherDataClusterHierarchyObjectsByDataClusterDef(aSuperiorDataClusters[i]);
        }
    }
}

/*---------------------------------------------------------------------------------------
    Check if export attribute "2b70adc0-4504-11e7-43b7-08002721906d" 
    (GDPR Processing Activity) is maintained and "true".
 ---------------------------------------------------------------------------------------*/   
function isGDPRRelevant(p_oItem) {
    try {
        var typeNum = ArisData.ActiveFilter().UserDefinedAttributeTypeNum("2b70adc0-4504-11e7-43b7-08002721906d");
        return isboolattributetrue(p_oItem, typeNum, g_nLoc);
    } catch (e) {
        return false;
    }
}

/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    ~~~~~~~ Common methods ~~~~~~~
  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/

/*---------------------------------------------------------------------------------------
    Combines all global Hashsets to a single one.
 ---------------------------------------------------------------------------------------*/
function combineGlobalHashSets() {
    var bigHashSet = new java.util.LinkedHashSet();
    
    //GDPR
    bigHashSet.addAll(g_dataClusterSet);
    
    //Test management
    bigHashSet.addAll(g_riskSet);
    bigHashSet.addAll(g_controlSet);
    bigHashSet.addAll(g_controlExecutionTaskSet);
    bigHashSet.addAll(g_testdefinitionSet);
    bigHashSet.addAll(g_userGroupSet);
    bigHashSet.addAll(g_userSet);
    
    bigHashSet.addAll(g_testerSet);
    bigHashSet.addAll(g_regulationSet);
    bigHashSet.addAll(g_orgUnitSet);
    bigHashSet.addAll(g_processSet);
    bigHashSet.addAll(g_riskCategorySet);
    
    //Risk management
    bigHashSet.addAll(g_appSysSet);
    
    //Survey management
    bigHashSet.addAll(g_questionnaireTemplateSet);
    bigHashSet.addAll(g_sectionSet);
    bigHashSet.addAll(g_questionSet);
    bigHashSet.addAll(g_optionSetSet);
    bigHashSet.addAll(g_optionSet);
    bigHashSet.addAll(g_surveyTaskSet);
	
	//Audit management
	bigHashSet.addAll(g_auditTemplateSet);
	bigHashSet.addAll(g_auditStepTemplateSet);
	
    //Policy management
    bigHashSet.addAll(g_policyDefinitionSet);
    
    return bigHashSet;
}

/*---------------------------------------------------------------------------------------
    Check if export attribute "AT_AAM_EXPORT_RELEVANT" is maintained and "true".
 ---------------------------------------------------------------------------------------*/   
function isExportRelevant(p_oItem) {
    return isboolattributetrue(p_oItem, Constants.AT_AAM_EXPORT_RELEVANT, g_nLoc);
}

/*---------------------------------------------------------------------------------------
    Checks if attribute "Regulations" is maintained and "true"
 ---------------------------------------------------------------------------------------*/   
function isRegulationItem(p_objDef) {
    return isboolattributetrue(p_objDef, Constants.AT_AAM_ANNUAL_ACCOUNTS_ITEM, g_nLoc);
}

/*---------------------------------------------------------------------------------------
    ---- CBA ----
	Checks if a given function is a control by default icon "control"
 ---------------------------------------------------------------------------------------*/   
function isControl(p_objDef) {
    var bIsFunction = p_objDef.TypeNum() == Constants.OT_FUNC;
    var bHasControlDefaultSymbol = p_objDef.GetDefaultSymbolNum() == Constants.ST_CONTR;
    
    return bIsFunction && bHasControlDefaultSymbol;
}

/*---------------------------------------------------------------------------------------
    Checks if the link between two given ObjDef is deactivated. The connection's direction
    does not matter.
    Returns either the boolean "true" or an empty String if the link is not deactivated
 ---------------------------------------------------------------------------------------*/
function isLinkDeactivated(objDef1, objDef2) {
    var bLinkDeactivated = g_deactivatedLinkSet.contains(objDef1.GUID() + "|" + objDef2.GUID())
                           || g_deactivatedLinkSet.contains(objDef2.GUID() + "|" + objDef1.GUID());
    if (bLinkDeactivated) {
        return bLinkDeactivated;
    } else {
        return "";
    }
}

/*---------------------------------------------------------------------------------------
    Stores the info if the link between two given ObjDefs is deactivated
 ---------------------------------------------------------------------------------------*/
function storeLinkDeactivated(objDef1, objDef2) {
    g_deactivatedLinkSet.add(objDef1.GUID() + "|" + objDef2.GUID());
}

/*---------------------------------------------------------------------------------------
    Loads the XML role mapping from the mapping file
    Convenience funtion for readArisRoleNameFromMapping()
 ---------------------------------------------------------------------------------------*/
function readRoleMapping() {   
    if (roleMappingXMLElement == null) {
        initConfigReader();
        if (g_xmlConfigReader != null && g_xmlConfigReader.isValid()) {
            var aEnumMappingsIterator = g_xmlConfigReader.getRootElement().getChild("aris2aamMapping").getChild("mappingEnums").getChildren("mappingEnum").iterator();
            while (aEnumMappingsIterator.hasNext()) {
                var enumMapping = aEnumMappingsIterator.next();
                if (enumMapping.getAttributeValue("enum").equals("userrole_type")) {
                    roleMappingXMLElement = enumMapping;
                    break;
                }
            }
        }
    }    
}
/*---------------------------------------------------------------------------------------
    Reads the the ARIS string for the given enum item value from the mapping file
    
    Example in standard mapping: for item value "clientadmin" it returns 
    "Client administrator"
    Convenience funtion for hasARCMRole()
 ---------------------------------------------------------------------------------------*/
function readArisRoleNameFromMapping(sEnumItemValue) {
    
    var sArisRoleName = null;
    readRoleMapping();
    if (roleMappingXMLElement != null) {
        var itemIterator = roleMappingXMLElement.getChildren("enumItem").iterator();
        while (itemIterator.hasNext()) {
            var roleItem = itemIterator.next();
            if (roleItem.getAttributeValue("id").equals(sEnumItemValue)) {
                sArisRoleName = roleItem.getAttributeValue("aris_typenum");
                break;
            }
        }
    }
    return sArisRoleName;
}
/*---------------------------------------------------------------------------------------
    Checks if a given person type p_userGroupDef represents a usergroups with the role 
    specified by param sARCMRole. If so then the funtion returns true, otherwise false.
    
    Example in standard mapping: if sARCMRole is "clientadmin" then the function returns 
    true if a person type is connected whose name starts with "Client administrator"
 ---------------------------------------------------------------------------------------*/
function hasARCMRole(p_userGroupDef, sARCMRole) {
    var bHasRole = false;
    //determine role
    var connectedRoles = getConnectedObjects(p_userGroupDef, Constants.EDGES_IN, Constants.CT_GENERAL_2, Constants.OT_PERS_TYPE);        
    var sMappedRoleName = readArisRoleNameFromMapping(sARCMRole);
    
    if (connectedRoles.length == 1 && sMappedRoleName != null) {       
        bHasRole =  connectedRoles[0].Name(g_nLoc).trim().startsWith(sMappedRoleName)
                    || connectedRoles[0].Name(1033).trim().startsWith(sMappedRoleName)
    }
    return bHasRole;
}

/*---------------------------------------------------------------------------------------
    Determines the objects that are linked to the given object considering the 
    CxnKind, CxnTypeNum, ObjTypeNum and SymbolTypeNum of the linked object.
	ObjTypeNum and SymbolTypeNum are optional - if not needed then they can be null or -1.
 ---------------------------------------------------------------------------------------*/   
function getConnectedObjects(p_objDef, p_cxnKind, p_cxnTypeNum, p_connectObjTypeNum, p_connectSymbolTypeNum) {
    var aConnectObjDefs = new Array();   
    
    var oCxns = null;
	if (p_cxnTypeNum != -1) {
		oCxns = p_objDef.CxnListFilter(p_cxnKind, p_cxnTypeNum);
	} else {
		oCxns = p_objDef.CxnListFilter(p_cxnKind);
	}
	
	//if a symbol is specified then find out which standard or user-defined symbol it is
	var iSymbolNum = null;
    if (p_connectSymbolTypeNum != null && isNaN(p_connectSymbolTypeNum)) {
        try {
            iSymbolNum = eval("Constants." + p_connectSymbolTypeNum );
        }
        catch (e) {
            iSymbolNum = eval("ArisData.ActiveFilter().UserDefinedSymbolTypeNum('" + p_connectSymbolTypeNum + "')" );
        }
    } else {
        if (p_connectSymbolTypeNum != -1) {
            iSymbolNum = p_connectSymbolTypeNum;
        } 
    }
		
    for (var i = 0; i < oCxns.length; i++ ) {
        var oConnectObjDef;
        if ( oCxns[i].SourceObjDef().equals(p_objDef) ) {
            oConnectObjDef = oCxns[i].TargetObjDef();
        } else {
            oConnectObjDef = oCxns[i].SourceObjDef();                
        }
        
        if (oConnectObjDef.GUID() == p_objDef.GUID()) {
            continue; //ignore cycle connections to itself
        }
		
		if (iSymbolNum != null && oConnectObjDef.getDefaultSymbolNum() != iSymbolNum) {
			continue;
		}
		
		if (p_connectObjTypeNum != null && p_connectObjTypeNum != -1 && oConnectObjDef.TypeNum() != p_connectObjTypeNum) {
			continue;
		}
		
		//ObjDef is qualified
        aConnectObjDefs.push(oConnectObjDef);
		if ( isboolattributetrue(oCxns[i], Constants.AT_DEACT, g_nLoc) ) {
			storeLinkDeactivated(p_objDef, oConnectObjDef);
		}       
        
    }    
    return aConnectObjDefs;
}

/*---------------------------------------------------------------------------------------
    As "getConnectedObjects()" - 
    Uses an Array of given ObjTypeNums though and therefore returns the linked objects
    for all these ObjTypeNums and SymbolTypeNum.
	ObjTypeNums and SymbolTypeNum are optional - if not needed then they can be null.
 ---------------------------------------------------------------------------------------*/   
function getConnectedObjects_A(p_objDef, p_cxnKind, p_cxnTypeNum, p_aConnectObjTypeNums, p_aConnectSymbolTypeNum) {
    
	if (p_aConnectObjTypeNums == null && p_aConnectSymbolTypeNum == null) {
		return getConnectedObjects(p_objDef, p_cxnKind, p_cxnTypeNum, null, null);
	}
	
	var aConnectObjDefs = new Array();
	if (p_aConnectObjTypeNums != null && p_aConnectSymbolTypeNum == null) {
		for (var i = 0; i < p_aConnectObjTypeNums.length; i++ ) {
			var currObjTypeNum = p_aConnectObjTypeNums[i];
			aConnectObjDefs = aConnectObjDefs.concat(getConnectedObjects(p_objDef, p_cxnKind, p_cxnTypeNum, currObjTypeNum, null));
		}
	}
	
	if (p_aConnectObjTypeNums == null && p_aConnectSymbolTypeNum != null) {
		for (var i = 0; i < p_aConnectSymbolTypeNum.length; i++ ) {
			var currSymbolTypeNum = p_aConnectSymbolTypeNum[i];
			aConnectObjDefs = aConnectObjDefs.concat(getConnectedObjects(p_objDef, p_cxnKind, p_cxnTypeNum, null, currSymbolTypeNum));
		}
	}

	if (p_aConnectObjTypeNums != null && p_aConnectSymbolTypeNum != null) {
		for (var i = 0; i < p_aConnectObjTypeNums.length; i++ ) {
			var currObjTypeNum = p_aConnectObjTypeNums[i];
			var currSymbolTypeNum = p_aConnectSymbolTypeNum[i];
			aConnectObjDefs = aConnectObjDefs.concat(getConnectedObjects(p_objDef, p_cxnKind, p_cxnTypeNum, currObjTypeNum, currSymbolTypeNum));
		}
	}
	 
    return aConnectObjDefs;
}

/*---------------------------------------------------------------------------------------
    Determines the role: the 1st object with a connection o type "is superior" 
 ---------------------------------------------------------------------------------------*/        
function getRoleObject(p_objDef) {
    if (p_objDef.TypeNum() == Constants.OT_PERS_TYPE) {
        var oCxns = p_objDef.CxnListFilter(Constants.EDGES_IN, Constants.CT_GENERAL_2);
        if (oCxns.length > 0) {
            return oCxns[0].SourceObjDef();
        }
    }
    return null;
}

/*---------------------------------------------------------------------------------------
    Classifies all ARIS person objects that are linked to one of the classified user groups
    by connection "CT_EXEC_5" as users
 ---------------------------------------------------------------------------------------*/  
function classifyUsers() {    
    for (var it = g_userGroupSet.iterator(); it.hasNext(); ) {
        var p_userGroupDef = it.next();
        var a_connectedUserDefs = getConnectedObjects(p_userGroupDef, Constants.EDGES_IN, Constants.CT_EXEC_5, Constants.OT_PERS);
        for (var i=0; i<a_connectedUserDefs.length; i++ ) {
            g_userSet.add(a_connectedUserDefs[i]);    
        }
    }
}

/*-----------------------------------------------------------------------------------------
    Updates the HashMaps which store the parent/child relations of the hierarchy elements.
 -----------------------------------------------------------------------------------------*/  
function updateHierarchyStructureHashMaps(p_hierarchy, p_aSuperiorHierarchies) {
    
    var parentsSet = convertJSArrayToHashSet(p_aSuperiorHierarchies);
    g_hierarchyChild2hierarchyParentsHashMap.put(p_hierarchy, parentsSet);
    
    for (var i=0; i<p_aSuperiorHierarchies.length ;i++) {
        var childrenArray = g_hierarchyParent2hierarchyChildrenHashMap.get(p_aSuperiorHierarchies[i]);
        if (childrenArray == null) {
            childrenArray = new Packages.java.util.HashSet(); 
            g_hierarchyParent2hierarchyChildrenHashMap.put(p_aSuperiorHierarchies[i], childrenArray);
        }
        childrenArray.add(p_hierarchy);
    }
}

/*-----------------------------------------------------------------------------------------
    Convenience method to check if a given ObjDef is contained in a given JSArray or Set.
 -----------------------------------------------------------------------------------------*/
function isObjDefContained(oObjDef, oCollection) {
    if (oObjDef == null || oCollection == null) {return false;}
    //JSArray
    if (oCollection.class == null) {
        for (var i=0; i<oCollection.length; i++) {
            if (oCollection[i].GUID() == oObjDef.GUID()) {return true;}
        }
        return false;
    }
    //java Set of some sort
    if (oCollection.class.toString().contains("Set")) {
        var it = oCollection.iterator();
        while (it.hasNext()) {
            var element = it.next();
            if (element.GUID() == oObjDef.GUID()) {return true;}
        }
    }
    return false;
}

/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    ~~~~~~~ Common methods ~~~~~~~
                    ~~~~~~~~~ Hierarchies ~~~~~~~~
  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
  
/*---------------------------------------------------------------------------------------
    Determines the objects that are linked to the given function:
    - Auditor groups / Sign-off owner groups
    - Superior function
 ---------------------------------------------------------------------------------------*/   
function getFurtherObjectsByFunc(p_oFuncDef) {
    //Function -> Auditor groups / Sign-off owner groups
    getUserGroupsAssignedToHierarchyElement(p_oFuncDef, true, false, true, false);
    getFurtherHierarchyObjects(p_oFuncDef, Constants.EDGES_IN | Constants.EDGES_ASSIGN, Constants.CT_IS_PRCS_ORNT_SUPER, Constants.OT_FUNC, g_processSet, null);
}

/*---------------------------------------------------------------------------------------
    Determines the objects that are linked to the given data cluster:
    - data cluster
 ---------------------------------------------------------------------------------------*/  
function getFurtherObjectsByData(p_oDataDef) {
    getUserGroupsAssignedToHierarchyElement(p_oDataDef, false, false, false, false);
    getFurtherHierarchyObjects(p_oDataDef, Constants.EDGES_IN | Constants.EDGES_ASSIGN, Constants.CT_CONS_OF_2, Constants.OT_CLST, g_dataClusterSet, null);
}

/*---------------------------------------------------------------------------------------
    Determines the objects that are linked to the given risk category:
	- Auditor groups
    - Superior risk category
 ---------------------------------------------------------------------------------------*/  
function getFurtherObjectsByRiskCategory(p_oRiskCategoryDef) {
    //Risk Category -> Auditor groups
    getUserGroupsAssignedToHierarchyElement(p_oRiskCategoryDef, true, false, false, false);
    getFurtherHierarchyObjects(p_oRiskCategoryDef, Constants.EDGES_IN, Constants.CT_BELONG_CAT, Constants.OT_RISK_CATEGORY, g_riskCategorySet, null);        
}

/*---------------------------------------------------------------------------------------
    Determines the objects that are linked to the given application system type:
	- Auditor groups
    - Superior application system type
 ---------------------------------------------------------------------------------------*/  
function getFurtherObjectsByAppSysType(p_oAppSysTypeDef) {
    //Application system type -> Auditor groups
    getUserGroupsAssignedToHierarchyElement(p_oAppSysTypeDef, true, false, false, false);
    getFurtherHierarchyObjects(p_oAppSysTypeDef, Constants.EDGES_IN, Constants.CT_CAN_SUBS_2, Constants.OT_APPL_SYS_TYPE, g_appSysSet, null);
}

/*---------------------------------------------------------------------------------------
    Determines the objects that are linked to the given organisational unit:
    - Auditor groups / Sign-off owner groups
    - Superior organisational unit (recursively)
    All superior organisational units are classified as belonging to the orgunit
    hierarchy.
 ---------------------------------------------------------------------------------------*/
function getFurtherOrgUnitHierarchyObjectsByOrgUnitDef(p_oOrgUnitDef) {
    //Organisational unit -> Auditor groups / Sign-off owner groups
    getUserGroupsAssignedToHierarchyElement(p_oOrgUnitDef, true, false, true, false);
    getFurtherHierarchyObjects(p_oOrgUnitDef, Constants.EDGES_IN, Constants.CT_IS_SUPERIOR_1, Constants.OT_ORG_UNIT, g_orgUnitSet, null);
}

/*---------------------------------------------------------------------------------------
    Determines the objects that are linked to the given organisational unit:
    - Auditor groups / Sign-off owner groups / Tester groups
    - Superior organisational unit (recursively)
    All superior organisational units are classified as belonging to the tester
    hierarchy.
 ---------------------------------------------------------------------------------------*/
function getFurtherTesterHierarchyObjectsByOrgUnitDef(p_oOrgUnitDef) {
    //Organisational unit -> Auditor groups / Sign-off owner groups / Tester groups
    getUserGroupsAssignedToHierarchyElement(p_oOrgUnitDef, true, false, true, true);
    getFurtherHierarchyObjects(p_oOrgUnitDef, Constants.EDGES_IN, Constants.CT_IS_SUPERIOR_1, Constants.OT_ORG_UNIT, g_testerSet, null);
}

/*---------------------------------------------------------------------------------------
    Determines the objects that are linked to the given regulation:
    - Auditor groups / Sign-off owner groups
    - Superior regulation
 ---------------------------------------------------------------------------------------*/
function getFurtherObjectsByRegulation(p_oRegulationDef) {
    //Regulation -> Auditor groups / Hierarchy owner groups / Sign-off owner groups
    getUserGroupsAssignedToHierarchyElement(p_oRegulationDef, true, true, true, false);
    getFurtherHierarchyObjects(p_oRegulationDef, Constants.EDGES_IN, Constants.CT_HAS_1, Constants.OT_TECH_TRM, g_regulationSet, Constants.AT_AAM_ANNUAL_ACCOUNTS_ITEM);
}

/*---------------------------------------------------------------------------------------
    Looks up hierarchy items in recursion.
    A boolean attribute can be specified; if so then its value must be "true", otherwise
    the superior hierarchy item is ignored.
 ---------------------------------------------------------------------------------------*/
function getFurtherHierarchyObjects(p_oHierarchyObjDef, p_iCxnDirection, p_iCxnType, p_iTargetObjDefType, p_oHandledHierarchiesSet, p_iQualificationBooleanAttribute) {	
    // Hierarchy item --> Superior hierarchy item
    if(!p_oHandledHierarchiesSet.add(p_oHierarchyObjDef)) {
        return;    
    }

    var aSuperiorQualifiedHierarchyDefs = new Array();
	var aSuperiorHierarchyDefs = getConnectedObjects(p_oHierarchyObjDef, p_iCxnDirection, p_iCxnType, p_iTargetObjDefType);
	for (var i=0; i<aSuperiorHierarchyDefs.length; i++) {
        if (p_iQualificationBooleanAttribute == null 
            || isboolattributetrue(aSuperiorHierarchyDefs[i], p_iQualificationBooleanAttribute, g_nLoc)
            || hasOccurrenceInQualifiedModel(aSuperiorHierarchyDefs[i], p_iQualificationBooleanAttribute)) {
            aSuperiorQualifiedHierarchyDefs.push(aSuperiorHierarchyDefs[i]);
        }
    }
    updateHierarchyStructureHashMaps(p_oHierarchyObjDef, aSuperiorQualifiedHierarchyDefs);
	//recursion
	for (var i=0; i<aSuperiorQualifiedHierarchyDefs.length; i++) {
        //if superior hierarchy item was not handled before
        getFurtherHierarchyObjects(aSuperiorQualifiedHierarchyDefs[i], p_iCxnDirection, p_iCxnType, p_iTargetObjDefType, p_oHandledHierarchiesSet, p_iQualificationBooleanAttribute);
    }
}

/*---------------------------------------------------------------------------------------
    Checks if the given ObjDef has an occurrence it at least one model where the
    given boolean attribute is set to 'true'.
 ---------------------------------------------------------------------------------------*/
function hasOccurrenceInQualifiedModel(p_oObjDef, p_iQualificationBooleanAttribute) {
    var objOccs = p_oObjDef.OccList();
    for (var i=0; i<objOccs.length; i++) {
        var oModel = objOccs[i].Model();    
        if (isboolattributetrue(oModel, p_iQualificationBooleanAttribute, g_nLoc)) {
            return true;
        }
    }
    return false;
}

/*---------------------------------------------------------------------------------------
    Remove tester from org unit set
 ---------------------------------------------------------------------------------------*/
function removeEntriesFromUrgUnitSetIfinTesterSet() {
    g_orgUnitSet.removeAll(g_testerSet);
}

/*---------------------------------------------------------------------------------------
    Looks up for  processes from model.
    - p_oModel model that is checked
    - aObjectsToClassify    
 ---------------------------------------------------------------------------------------*/
function getProcessesFromModel(p_oModel, aObjectsToClassify) {
    if(g_processModelClassified.add(p_oModel)) {
        buildAllowedModels();
        // processes    
        if(isProcessModelAllowed(p_oModel)) {            
            for (var j = 0; j < aObjectsToClassify.length; j++) {
                // Func --> ... 
                var oFuncDef = aObjectsToClassify[j];
                // add current process to global set
                // if process was not handled before
                if(!isControl(oFuncDef) && isProcessObjectAllowed(oFuncDef) && isProcessSymbolAllowed(p_oModel, oFuncDef)) {
                    getFurtherObjectsByFunc(oFuncDef);
                }
            }
        }   
    }    
}

function isProcessModelAllowed(oModel) {
    return isModelAllowed(PROCESS_MODEL, oModel);
}

function isProcessObjectAllowed(oFuncDef) {
    return isObjectAllowed(PROCESS_MODEL, oFuncDef);
}

function isProcessSymbolAllowed(oModel, oFuncDef) {    
    return isSymbolAllowed(PROCESS_MODEL, oModel, oFuncDef);
}

/*---------------------------------------------------------------------------------------
    Looks up for data cluster from model.
    - p_oModel model that is checked
    - aObjectsToClassify    
 ---------------------------------------------------------------------------------------*/
function getDataClusterFromModel(p_oModel, aObjectsToClassify) {
    if(g_dataClusterModelClassified.add(p_oModel)) {
        buildAllowedModels();
        if(isDataClusterModelAllowed(p_oModel)) {            
            for (var i = 0; i < aObjectsToClassify.length; i++) {
                // Cluster --> ...
                var oDataDef = aObjectsToClassify[i];
                if(isDataClusterObjectAllowed(oDataDef) && isDataClusterSymbolAllowed(p_oModel, oDataDef) ) {
                    // add current process to global set
                    // if process was not handled before
                    getFurtherObjectsByData(oDataDef);
                }
            }
        }   
    }    
}

function isDataClusterModelAllowed(oModel) {
    return isModelAllowed(DATA_CLUSTER_MODEL, oModel);
}

function isDataClusterObjectAllowed(oDataDef) {
    return isObjectAllowed(DATA_CLUSTER_MODEL, oDataDef);
}

function isDataClusterSymbolAllowed(p_oModel, oDataDef) {    
    return isSymbolAllowed(DATA_CLUSTER_MODEL, p_oModel, oDataDef);
}

/*---------------------------------------------------------------------------------------
    Looks up for processes from model.
    - p_oModel model that is checked
    - aObjectsToClassify    
 ---------------------------------------------------------------------------------------*/
function getRegulationsFromModel(p_oModel, aObjectsToClassify) {
    if(g_regulationModelClassified.add(p_oModel)) {
        buildAllowedModels();
        if(isRegulationModelAllowed(p_oModel)) { 
            for (var i=0; i<aObjectsToClassify.length; i++) {
                //technical terms -> regulation hierarchy
                var oRegulationDef = aObjectsToClassify[i];
                if (isRegulationObjectAllowed(oRegulationDef) && isSymbolAllowed(p_oModel, oRegulationDef)
                    // remove later after method is changed
                    && (isRegulationItem(oRegulationDef) || isRegulationItem(p_oModel))
                    ) {
                    // add current regulations to global set
                    // if regulation was not handled before
                    getFurtherObjectsByRegulation(oRegulationDef);
                }
            }
        }  
    }
}

function isRegulationModelAllowed(oModel) {
    return isModelAllowed(REGULATION_MODEL, oModel);
}

function isRegulationObjectAllowed(oRegulationDef) {
    return isObjectAllowed(REGULATION_MODEL, oRegulationDef);
}

function isSymbolAllowed(oModel, oRegulationDef) {
    return isSymbolAllowed(REGULATION_MODEL, oModel, oRegulationDef);
}

/*---------------------------------------------------------------------------------------
    Looks up for application system types from model.
    - p_oModel model that is checked
    - aObjectsToClassify    
 ---------------------------------------------------------------------------------------*/
function getApplicationSystemTypesFromModel(p_oModel, aObjectsToClassify) {
    if(g_applicationSystemTypeModelClassified.add(p_oModel)) {
        buildAllowedModels();                
        if(isApplicationSystemTypeModelAllowed(p_oModel)) {
            for (var i=0; i<aObjectsToClassify.length; i++) {
                var oAppSysDef = aObjectsToClassify[i];
                if (isApplicationSystemTypeObjectAllowed(oAppSysDef) && isApplicationSystemTypeSymbolAllowed(p_oModel, oAppSysDef)) {
                    getFurtherObjectsByAppSysType(oAppSysDef);
                }
            }    
        }
    }
}

function isApplicationSystemTypeModelAllowed(oModel) {
    return isModelAllowed(APPLICATION_SYSTEM_TYPE_MODEL, oModel);
}

function isApplicationSystemTypeObjectAllowed(oAppSysDef) {
    return isObjectAllowed(APPLICATION_SYSTEM_TYPE_MODEL, oAppSysDef);
}

function isApplicationSystemTypeSymbolAllowed(oModel, oAppSysDef) {
    return isSymbolAllowed(APPLICATION_SYSTEM_TYPE_MODEL, oModel, oAppSysDef);    
}

/*---------------------------------------------------------------------------------------
    Looks up for  risk categories from models.
    - p_oModel model that is checked
    - aObjectsToClassify    
 ---------------------------------------------------------------------------------------*/
function getRiskCategoriesFromModel(p_oModel, aObjectsToClassify) {
	if(g_riskCategoryModelClassified.add(p_oModel)) {
        buildAllowedModels();
        if(isRiskCategoryModelAllowed(p_oModel)) {         
			for (var i=0; i<aObjectsToClassify.length; i++) {
				//questionnaire templates in this model
                var oRiskCatDef = aObjectsToClassify[i];
                if(isRiskCategoryObjectAllowed(oRiskCatDef) && isRiskCategorySymbolAllowed(p_oModel, oRiskCatDef)) {
					getFurtherObjectsByRiskCategory(oRiskCatDef);
				}
			}
		}
	}
}

function isRiskCategoryModelAllowed(oModel) {
    return isModelAllowed(RISK_CATEGORY_MODEL, oModel);
}

function isRiskCategoryObjectAllowed(oRiskCatDef) {
    return isObjectAllowed(RISK_CATEGORY_MODEL, oRiskCatDef);
}

function isRiskCategorySymbolAllowed(oModel, oRiskCatDef) {
    return isSymbolAllowed(RISK_CATEGORY_MODEL, oModel, oRiskCatDef);    
}

/*---------------------------------------------------------------------------------------
    Looks up for org units from models.
    - p_oModel model that is checked
    - aObjectsToClassify    
 ---------------------------------------------------------------------------------------*/
function getOrgUnitsFromModel(p_oModel, aObjectsToClassify)	{    
    if(g_orgUnitModelClassified.add(p_oModel)) {
        buildAllowedModels();                
        if(isOrgUnitModelAllowed(p_oModel)) {  
            for (var i=0; i<aObjectsToClassify.length; i++) {
                //ABA organisational unit objects
                var oOrgUnitDef = aObjectsToClassify[i];
                if (isOrgUnitObjectAllowed(oOrgUnitDef) && isOrgUnitSymbolAllowed(p_oModel, oOrgUnitDef)) {                
                    //orgunits -> organisational hierarchy
                    getFurtherOrgUnitHierarchyObjectsByOrgUnitDef(oOrgUnitDef);
                }
            }    
        }	
    }
}

function isOrgUnitModelAllowed(oModel) {
    return isModelAllowed(ORG_UNIT_MODEL, oModel);
}

function isOrgUnitObjectAllowed(oOrgUnitDef) {
    return isObjectAllowed(ORG_UNIT_MODEL, oOrgUnitDef);
}

function isOrgUnitSymbolAllowed(oModel, oOrgUnitDef) {
    return isSymbolAllowed(ORG_UNIT_MODEL, oModel, oOrgUnitDef);    
}

/*---------------------------------------------------------------------------------------
    build from aris2arcm-mapping.xml (allowedModels element in xml)
    key: e.g. PROCESS_MODEL
    value: set of model type numbers, set object type numbers, set of symbol type numbers 
 ---------------------------------------------------------------------------------------*/
function buildAllowedModels() {
    // init only once
        
    if(Object.keys(g_allowedModels).length == 0) {
        if(g_xmlConfigReader == null) {
            initConfigReader();
        }        
        var xmlConfig_Root = g_xmlConfigReader.getRootElement();
        var xmlAllowedModels = xmlConfig_Root.getChild("allowedModels");
        
        var iterModels = xmlAllowedModels.getChildren("model").iterator();
        
        while (iterModels.hasNext()) {
            var model = iterModels.next();
            addAllowedModels(model);
            addAllowedObjects(model);
            addAllowedSymbols(model);
        }    
    }
}

// build from aris2arcm-mapping.xml (allowedModels element in xml)
// key: e.g. PROCESS_MODEL
// value: set of model type numbers
function addAllowedModels(xmlAllowedModel) {
    var modelType = xmlAllowedModel.getAttribute("type").getValue();
    var modelTypesList = xmlAllowedModel.getAttribute("models");
    
    if(modelTypesList == null) {
        return;
    }   
    
    modelTypesList = modelTypesList.getValue().replaceAll(" ", "");
    var modelTypes = modelTypesList.split(",");
    var models = new java.util.HashSet();
    for(var i = 0; i < modelTypes.length; i++) {
        models.add(getModelTypeNum(modelTypes[i]));
    }
    g_allowedModels[modelType] = models;    
}

// build from aris2arcm-mapping.xml (allowedModels element in xml)
// key: e.g. PROCESS_MODEL
// value: set of model type numbers
function addAllowedObjects(xmlAllowedModel) {
    var modelType = xmlAllowedModel.getAttribute("type").getValue();
    var objectTypesList = xmlAllowedModel.getAttribute("objects");
    
    if(objectTypesList == null) {
        return;
    }    
    
    objectTypesList = objectTypesList.getValue().replaceAll(" ", "");
    var objectTypes = objectTypesList.split(",");
    var objects = new java.util.HashSet();
    for(var i = 0; i < objectTypes.length; i++) {
        objects.add(getObjectNum(objectTypes[i]));
    }
    g_allowedObjects[modelType] = objects;    
}


// build from aris2arcm-mapping.xml (allowedModels element in xml)
// key: e.g. PROCESS_MODEL
// value: set of symbol type numbers
function addAllowedSymbols(xmlAllowedModel) {
    var modelType = xmlAllowedModel.getAttribute("type").getValue();
    var symbolTypesList = xmlAllowedModel.getAttribute("symbols");
    
    if(symbolTypesList == null) {
        return;
    }
    
    symbolTypesList = symbolTypesList.getValue().replaceAll(" ", "");
    var symbolTypes = symbolTypesList.split(",");
    var symbols = new java.util.HashSet();
    for(var i = 0; i < symbolTypes.length; i++) {
        symbols.add(getSymbolNum(symbolTypes[i]));
    }
    g_allowedSymbols[modelType] = symbols;    
}


function isModelAllowed(type, oModel) {
    var allowedModelss = g_allowedModels[type];
    if(allowedModelss == null) {
        return true;
    } 
    return allowedModelss.contains(oModel.TypeNum());
}

function isObjectAllowed(type, oObjDef) {
    var allowedObjects = g_allowedObjects[type];
    if(allowedObjects == null) {
        return true;
    }    
    return allowedObjects.contains(oObjDef.TypeNum());
}

function isSymbolAllowed(type, oModel, oObjDef) {
    var allowedSymbols = g_allowedSymbols[type];
    if(allowedSymbols == null) {
        return true;
    }
    var oOccs = oObjDef.OccListInModel(oModel);
    for (var i in oOccs) {
        if(allowedSymbols.contains(oOccs[i].SymbolNum())) {
            return true;
        }
    }
    return false
}

/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
         ~~~~~~~ TypeNm helper functions ~~~~~~~
  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/

/*---------------------------------------------------------------------------------------
    Checks the attribute type number:
    - for empty values -1 is returned
    - numerical values are not changed but only returned
    - if they belong to standard items the value of "Constants." + <p_sTypeNum>
      is returned
    - if they belong to custom items with GUID the currently valid TypeNumber
      is returned.
    - if it does not belong to any category above -1 is returned
---------------------------------------------------------------------------------------*/
function obtainTypeNumFromFilter(p_sTypeNum, p_sFilterFunction) {
    
    var sTypeNum = new java.lang.String(p_sTypeNum);  
    
    //if no TypeNum was set
    if (sTypeNum.length() == 0) {return -1;}
    
    //if TypeNum is already a number just return it
    if (!isNaN(sTypeNum)) {return Number(p_sTypeNum);}
 
    //if TypeNum belongs to a standard attribute return the corresponding Constant
    var iTypeNum;
    try {
        iTypeNum = eval("Constants." + sTypeNum);
        if (!isNaN(iTypeNum)) {
            return Number(iTypeNum);
        }
    }
    catch (e) {
        //ok, it does not
    }
 
    //check if TypeNum is the GUID of a user-defined attribute  
    try {
        if (p_sFilterFunction != null) {
            iTypeNum = eval("ArisData.ActiveFilter()." + p_sFilterFunction + "('" + sTypeNum.trim() + "')" );
            if (!isNaN(iTypeNum)) {
                return Number(iTypeNum);
            }
        }
    }
    catch (e) {
        //ok, it is not
    }
    
    //unknown TypeNum
    return -1;
}

function getAttributeTypeNum(p_sTypeNum) {
    return obtainTypeNumFromFilter(p_sTypeNum, "UserDefinedAttributeTypeNum");
}

function getModelTypeNum(p_sTypeNum) {
    return obtainTypeNumFromFilter(p_sTypeNum, "UserDefinedModelTypeNum");
}

function getSymbolNum(p_sSymbolNum) {
    return obtainTypeNumFromFilter(p_sSymbolNum, "UserDefinedSymbolTypeNum");
}

function getObjectNum(p_sObjectNum) {
    return obtainTypeNumFromFilter(p_sObjectNum, null);
}

/*---------------------------------------------------------------------------------------
    Returns 'undefined' instead of -1 if p_sTypeNum is unknown
---------------------------------------------------------------------------------------*/
function getAttributeTypeNum_Enum(p_sTypeNum) {
    var iTypeNum = obtainTypeNumFromFilter(p_sTypeNum, "UserDefinedAttributeTypeNum");
    if (iTypeNum == -1 || isNaN(iTypeNum)) {return undefined;} //HOTFIX SR5 - throws no exception anymore but returns -1
    return iTypeNum;
}