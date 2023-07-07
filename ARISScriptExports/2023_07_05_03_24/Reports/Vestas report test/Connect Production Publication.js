var nLocale = Context.getSelectedLanguage()

//SEARCH FOR MODELS TO PUBLISH
var design_DB = ArisData.openDatabase("Design", "system", "manager", "dd838074-ac29-11d4-85b8-00005a4053ff", nLocale, false)
var searchItemSpecial1 = design_DB.createSearchItem(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("adde7a20-75f9-11e0-61d8-00237d347a8a"), nLocale, "Yes", Constants.SEARCH_CMP_EQUAL, false, false)
var searchItemSpecial2 = design_DB.createSearchItem(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("8fda52b0-6188-11de-4732-00237d347a8a"), nLocale, "All attributes filled-in", Constants.SEARCH_CMP_EQUAL, false, false)
var searchItemSpecial3 = design_DB.createSearchItem(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("27204be0-fb11-11e2-61d8-00237d347a8a"), nLocale, true)

//FIND Value Chain/Decision Model
//FIND Organizational Chart
//FIND Process Flow Chart – Standard
//FIND Function Tree
//FIND Function Allocation diagram
//FIND Information Carrier Diagram
//FIND Structuring model
var mixModelsForPublication = design_DB.Find(Constants.SEARCH_MODEL, [Constants.MT_VAL_ADD_CHN_DGM, Constants.MT_ORG_CHRT, Constants.MT_EEPC, Constants.MT_FUNC_TREE, Constants.MT_FUNC_ALLOC_DGM, Constants.MT_INFO_CARR_DGM, Constants.MT_STRCT_DGM], searchItemSpecial1)

//SPECIAL FIND Process Flowchart Swimlanes
var processFlowchartSwimlanes = design_DB.Find(Constants.SEARCH_MODEL, [Constants.MT_EEPC_COLUMN], searchItemSpecial1.and(searchItemSpecial2.and(searchItemSpecial3)))

//**********Concatenate models*********
/* var modelsForPublication = new Array()

for (var i = 0; i < mixModelsForPublication.length; i++) {
    modelsForPublication.push(mixModelsForPublication[i])
    mixModelsForPublication[i] = null;
}
mixModelsForPublication = null;

for (var i = 0; i < processFlowchartSwimlanes.length; i++) {
    modelsForPublication.push(processFlowchartSwimlanes[i])
    processFlowchartSwimlanes[i] = null;
}
processFlowchartSwimlanes = null; */
//*************************************

//Set attributes values in mixModelsForPublication as released
for (var i = 0; i < mixModelsForPublication.length; i++) {
    var modelReadyForVppPublishValue_Released = ArisData.ActiveFilter().UserDefinedAttributeValueTypeNum("adde7a20-75f9-11e0-61d8-00237d347a8a", "adde7a22-75f9-11e0-61d8-00237d347a8a")
    var modelReadyForVppPublishWrite = mixModelsForPublication[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("adde7a20-75f9-11e0-61d8-00237d347a8a"), nLocale).setValue(modelReadyForVppPublishValue_Released)
    var statusValue_Released = ArisData.ActiveFilter().UserDefinedAttributeValueTypeNum("8fda52b0-6188-11de-4732-00237d347a8a", "dc5a6e45-e572-11de-4732-00237d347a8a")
    var statusWrite = mixModelsForPublication[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("8fda52b0-6188-11de-4732-00237d347a8a"), nLocale).setValue(statusValue_Released)
}

//Set attributes values in processFlowchartSwimlanes as released
for (var i = 0; i < processFlowchartSwimlanes.length; i++) {
    var modelReadyForVppPublishValue_Released = ArisData.ActiveFilter().UserDefinedAttributeValueTypeNum("adde7a20-75f9-11e0-61d8-00237d347a8a", "adde7a22-75f9-11e0-61d8-00237d347a8a")
    var modelReadyForVppPublishWrite = processFlowchartSwimlanes[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("adde7a20-75f9-11e0-61d8-00237d347a8a"), nLocale).setValue(modelReadyForVppPublishValue_Released)
    var statusValue_Released = ArisData.ActiveFilter().UserDefinedAttributeValueTypeNum("8fda52b0-6188-11de-4732-00237d347a8a", "dc5a6e45-e572-11de-4732-00237d347a8a")
    var statusWrite = processFlowchartSwimlanes[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("8fda52b0-6188-11de-4732-00237d347a8a"), nLocale).setValue(statusValue_Released)
}

//*****************************************************
//MERGE MODELS TO Publishied DB
var published_DB = ArisData.openDatabase("Published", "system", "manager", "dd838074-ac29-11d4-85b8-00005a4053ff", nLocale, false)
var merge = Context.getComponent("Merge")
var mergeOptions = merge.createMergeOptions(design_DB)
mergeOptions.setAssignmentDepth(0) //***this MUST be zero (0)***
mergeOptions.setMixAttributes(false) //do not merge attributes
mergeOptions.setConsiderBorderAssignment(true)
mergeOptions.setConsiderBorderCxnDefs(true)
mergeOptions.setMixUserRelations(false) // Specifies whether relationships between users and user groups are to be merged.  
mergeOptions.setConsiderSourceGroupPath(true) // If this option is set, then the merge algorithm will use the source group path in the case of a conflict.
mergeOptions.setInheritAccessRightsForNewGroups(false) //true if access rights should be inherited when creating new groups
mergeOptions.setSynchronizeAssignments(true) //delete assignment on target side

var mergeResult = merge.startMerge(processFlowchartSwimlanes, published_DB, mergeOptions, true)
var mergeResult = merge.startMerge(mixModelsForPublication, published_DB, mergeOptions, true)

published_DB.clearCaches()
design_DB.clearCaches()
//*****************************************************

//***********************************************************************
//DELETE OBSOLETE MODELS
var searchItem1 = design_DB.createSearchItem(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("adde7a20-75f9-11e0-61d8-00237d347a8a"), nLocale, "Process Obsolete", Constants.SEARCH_CMP_EQUAL, false, false)
var searchItem2 = design_DB.createSearchItem(Constants.AT_REM, nLocale, false)

var obsoleteModelsDesign = design_DB.Find(Constants.SEARCH_MODEL, null, searchItem1.and(searchItem2))

for (var i = 0; i < obsoleteModelsDesign.length; i++) {

    var modelIdentifier = obsoleteModelsDesign[i].Attribute(Constants.AT_ID, nLocale).getValue()
    var searchItem3 = design_DB.createSearchItem(Constants.AT_ID, nLocale, modelIdentifier, Constants.SEARCH_CMP_EQUAL, false, false)
    var searchItem4 = design_DB.createSearchItem(Constants.AT_NAME, nLocale, obsoleteModelsDesign[i].Name(nLocale), Constants.SEARCH_CMP_EQUAL, false, false)
    var obsoleteModelsReview = published_DB.Find(Constants.SEARCH_MODEL, null, searchItem3.and(searchItem4))

    //delete models from Review
    deleteModels(obsoleteModelsReview)

    //change att values in Design
    var statusValue_ProcessObsolete = ArisData.ActiveFilter().UserDefinedAttributeValueTypeNum("8fda52b0-6188-11de-4732-00237d347a8a", "60e80820-0143-11df-4732-00237d347a8a")
    var statusWrite = obsoleteModelsDesign[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("8fda52b0-6188-11de-4732-00237d347a8a"), nLocale).setValue(statusValue_ProcessObsolete)
    var remarkExampleWrite = obsoleteModelsDesign[i].Attribute(Constants.AT_REM, nLocale).setValue("NOT IN VPP")
}
published_DB.clearCaches()
//published_DB.close()
design_DB.clearCaches()
design_DB.close()
//***********************************************************************

//REORGANIZE DB
reorganizeDb(published_DB)
published_DB.close()

function deleteModels(models) {
    for (var i = 0; i < models.length; i++) {
        var modelGroup = models[i].Group();
        modelGroup.Delete(models[i]);
    }
}

function reorganizeDb(dbName) {
    var adminComponent = Context.getComponent("ServerAdmin")    
    var result = adminComponent.reorganiseArisDatabase(dbName, null)
}