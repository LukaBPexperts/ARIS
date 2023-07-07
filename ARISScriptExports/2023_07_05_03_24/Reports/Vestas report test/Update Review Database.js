var nLocale = Context.getSelectedLanguage()

var selectedModels = ArisData.getSelectedModels()

var publishModels = findModelsForPublish(selectedModels)
var obsoleteModels = findObsoleteModels(selectedModels)

var btnPressed = Dialogs.MsgBox(publishModels.length + " model will be published and " + obsoleteModels.length + " model will be removed from portal. Do you want to continue?", Constants.MSGBOX_ICON_QUESTION | Constants.MSGBOX_BTN_YESNO + 512, "Message from Report")
if (btnPressed == Constants.MSGBOX_RESULT_YES) {

    //Set attributes values in models
    for (var i = 0; i < publishModels.length; i++) {
        var modelReadyForTestValue_Released = ArisData.ActiveFilter().UserDefinedAttributeValueTypeNum("3cfde370-e44c-11e9-004f-eabd1a8bb6a2", "3cfde372-e44c-11e9-004f-eabd1a8bb6a2")
        var modelReadyForTestWrite = publishModels[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("3cfde370-e44c-11e9-004f-eabd1a8bb6a2"), nLocale).setValue(modelReadyForTestValue_Released)
        //var statusValue_Released = ArisData.ActiveFilter().UserDefinedAttributeValueTypeNum("8fda52b0-6188-11de-4732-00237d347a8a", "dc5a6e45-e572-11de-4732-00237d347a8a")
        //var statusWrite = publishModels[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("8fda52b0-6188-11de-4732-00237d347a8a"), nLocale).setValue(statusValue_Released)
    }

    //*****************************************************
    //MERGE MODELS TO Publishied DB
    var review_DB = ArisData.openDatabase("Review", "system", "manager", "dd838074-ac29-11d4-85b8-00005a4053ff", nLocale, false)
    var design_DB = ArisData.openDatabase("Design", "system", "manager", "dd838074-ac29-11d4-85b8-00005a4053ff", nLocale, false)
    var merge = Context.getComponent("Merge")
    var mergeOptions = merge.createMergeOptions(design_DB)
    mergeOptions.setAssignmentDepth(0) //***this MUST be zero (0)***
    mergeOptions.setMixAttributes(false) //do not merge attributes
    mergeOptions.setConsiderBorderAssignment(true)
    mergeOptions.setConsiderBorderCxnDefs(true)
    mergeOptions.setMixUserRelations(false) // Specifies whether relationships between users and user groups are to be merged.  
    mergeOptions.setConsiderSourceGroupPath(true) // If this option is set, then the merge algorithm will use the source group path in the case of a conflict.
    mergeOptions.setInheritAccessRightsForNewGroups(false) //true if access rights should be inherited when creating new groups
    //mergeOptions.setSynchronizeAssignments(true) //delete assignment on target side

    var mergeResult = merge.startMerge(publishModels, review_DB, mergeOptions, true)

    review_DB.clearCaches()
    design_DB.clearCaches()
    //*****************************************************

    //***********************************************************************
    //DELETE OBSOLETE MODELS
    for (var i = 0; i < obsoleteModels.length; i++) {

        var modelIdentifier = obsoleteModels[i].Attribute(Constants.AT_ID, nLocale).getValue()
        var searchItem3 = design_DB.createSearchItem(Constants.AT_ID, nLocale, modelIdentifier, Constants.SEARCH_CMP_EQUAL, false, false)
        var searchItem4 = design_DB.createSearchItem(Constants.AT_NAME, nLocale, obsoleteModels[i].Name(nLocale), Constants.SEARCH_CMP_EQUAL, false, false)
        var obsoleteModelsReview = review_DB.Find(Constants.SEARCH_MODEL, null, searchItem3.and(searchItem4))

        //delete models from Review
        deleteModels(obsoleteModelsReview)

        //change att values in Design
        //var statusValue_ProcessObsolete = ArisData.ActiveFilter().UserDefinedAttributeValueTypeNum("8fda52b0-6188-11de-4732-00237d347a8a", "60e80820-0143-11df-4732-00237d347a8a")
        //var statusWrite = obsoleteModels[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("8fda52b0-6188-11de-4732-00237d347a8a"), nLocale).setValue(statusValue_ProcessObsolete)
        //var remarkExampleWrite = obsoleteModels[i].Attribute(Constants.AT_REM, nLocale).setValue("NOT IN VPP")
    }
    review_DB.clearCaches()
    //review_DB.close()
    design_DB.clearCaches()
    design_DB.close()
    //***********************************************************************

    //REORGANIZE DB
    reorganizeDb(review_DB)
    review_DB.close()

}

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

function findModelsForPublish(selectedModels) {
    var models = new Array()
    for (var i = 0; i < selectedModels.length; i++) {
        var modelReadyForReviewPublishValue = selectedModels[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("3cfde370-e44c-11e9-004f-eabd1a8bb6a2"), nLocale).getValue()
        //var statusValue = selectedModels[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("8fda52b0-6188-11de-4732-00237d347a8a"), nLocale).getValue()
        //var changeTypeMaintained = selectedModels[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("27204be0-fb11-11e2-61d8-00237d347a8a"), nLocale).IsMaintained()

        if (selectedModels[i].TypeNum() == Constants.MT_EEPC_COLUMN) {
            if (modelReadyForReviewPublishValue == "Yes") {
                models.push(selectedModels[i])
            }
        } else {
            if (modelReadyForReviewPublishValue == "Yes") {
                models.push(selectedModels[i])
            }
        }
    }
    return models
}

function findObsoleteModels(selectedModels) {
    var models = new Array()
    for (var i = 0; i < selectedModels.length; i++) {
        var modelReadyForReviewPublishValue = selectedModels[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("3cfde370-e44c-11e9-004f-eabd1a8bb6a2"), nLocale).getValue()
        //var remarkExampleMaintained = selectedModels[i].Attribute(Constants.AT_REM, nLocale).IsMaintained()

        if (modelReadyForReviewPublishValue == "Process Obsolete") {
            models.push(selectedModels[i])
        }
    }
    return models
}