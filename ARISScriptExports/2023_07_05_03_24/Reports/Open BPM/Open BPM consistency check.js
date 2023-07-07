/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

var gb_c_INFO_MARK_SET      = false;
var gb_c_WARNING_MARK_SET   = false;
var gb_c_ERROR_MARK_SET     = false;
var gb_DGRMS_OPEN_SET       = false;

INFO    = function(sErrors, sWarnings, sInfos, nType) {
    this.sErrors    = sErrors;
    this.sWarnings  = sWarnings;
    this.sInfos     = sInfos;
    this.nType      = nType;
};

var g_semCheckComponent = Context.getComponent("BPMNSemanticChecks");

var gn_Lang         = Context.getSelectedLanguage();
var g_oDB           = ArisData.getActiveDatabase();
var gs_Empty        = new java.lang.String("");
var gs_cEmpty       = new java.lang.String("-");
var gn_Err          = -1;
var BPMN2XML        = false;

/*************************************************************************************************************************************/

function check_ModelForErrors( mapInfoMarks, aModels ){
    for(var i=0; i< aModels.length; i++){
        var bCycle = false;
        check_ManualTask( mapInfoMarks, aModels[i] );
        check_UserTask( mapInfoMarks, aModels[i] );
        check_ServiceTask( mapInfoMarks, aModels[i] );
        if (BPMN2XML === false) check_AbstractTask( mapInfoMarks, aModels[i] );
        check_SpecialTaskRelations( mapInfoMarks, aModels[i] );

        check_SyncAgainstCentraSite( mapInfoMarks, aModels[i] );
        check_AssignedFadModel( mapInfoMarks, aModels[i] );
        bCycle = checkTaskHierarchyCycle( mapInfoMarks, aModels[i] );
        checkTaskAssignment( mapInfoMarks, aModels[i] );

        check_wMChecks( mapInfoMarks, aModels[i] );

        var aSubProcessModels   = getAssignedSubProcessModels( aModels[i] );
        if( !bCycle && aSubProcessModels.length > 0 ) {
            check_ModelForErrors( mapInfoMarks, aSubProcessModels );
        }
    }
}

function getAssignedSubProcessModels(p_oModel) {
    var aOccs   = getSubprocessesInModel(p_oModel);
    //aOccs = aOccs.concat(getCallActivitiesInModel(p_oModel));//ARHI: Recursion must also include assignments of call activities
    //ARHI: As call activities are only referenced, the recursion isn't needed anymore
    var aAssignedModels = new Array();
    for(var i=0; i<aOccs.length; i++){
        aAssignedModels = aAssignedModels.concat(getAssignedModelsOfBasType(aOccs[i].ObjDef(), [Constants.MT_BPMN_COLLABORATION_DIAGRAM, Constants.MT_BPMN_PROCESS_DIAGRAM]));
    }
    aAssignedModels.clearDuplicities();
    return aAssignedModels;
}

function check_ManualTask(mapInfoMarks, oModel) {
    var oManualTaskOccs = getManualTasksInModel(oModel);
    for (var i = 0; i < oManualTaskOccs.length; i++) {
        var oManualTaskOcc      = oManualTaskOccs[i];

        // Manual Task is carried out by organizational element
        var oCxns2OrgElements   = getCxns2OrgElements(oManualTaskOcc);
        var oCxns2SSTs          = getCxns2SSTs(oManualTaskOcc);
        var oCxns2SSOTs         = getCxns2SSOTs(oManualTaskOcc);

        if (oCxns2OrgElements.length > 1 ||
            (BPMN2XML === false && !(oCxns2OrgElements.length == 1 && oCxns2SSTs.length == 0 && oCxns2SSOTs.length == 0))) {
            setInfoMark(mapInfoMarks, oManualTaskOcc, getString("WRG_MANUAL_TASK"), Constants.MODEL_INFO_WARNING);
        }
    }
}

function check_UserTask(mapInfoMarks, oModel) {
    var oUserTaskOccs = getUserTasksInModel(oModel);
    for (var i = 0; i < oUserTaskOccs.length; i++) {
        var oUserTaskOcc = oUserTaskOccs[i];
        var oCxns2OrgElements = getCxns2OrgElements(oUserTaskOcc);
        var numberOfOrgElements = oCxns2OrgElements.length;

        if (BPMN2XML === false) {
            var oCxns2SSTs = getCxns2SSTs(oUserTaskOcc);
            var oCxns2SSOTs = getCxns2SSOTs(oUserTaskOcc);
            var numberOfServices = oCxns2SSTs.length;
            var numberOfServiceOperations = oCxns2SSOTs.length;

            if (checkISService(oCxns2SSTs)) {
                setInfoMark(mapInfoMarks, oUserTaskOcc, getString("ERR_IS_SERVICE_USER_TASK"), Constants.MODEL_INFO_ERROR);
            } else if (numberOfOrgElements > 1 || numberOfServices > 1 || numberOfServiceOperations > 1) {
                setInfoMark(mapInfoMarks, oUserTaskOcc, getString("WRG_USER_TASK"), Constants.MODEL_INFO_WARNING);
            } else if (numberOfOrgElements == 1) {
                if ((numberOfServices + numberOfServiceOperations) > 0) {
                    setInfoMark(mapInfoMarks, oUserTaskOcc, getString("WRG_USER_TASK"), Constants.MODEL_INFO_WARNING);
                }
            } else {
                if (numberOfServices == 1 && numberOfServiceOperations == 1) {
                    if (!checkSSOT2SST_Relation(oCxns2SSOTs, oCxns2SSTs)) {
                        setInfoMark(mapInfoMarks, oUserTaskOcc, getString("ERR_SST_SSOT_NOT_RELATED"), Constants.MODEL_INFO_ERROR);
                    }
                } else if ((numberOfServices + numberOfServiceOperations) > 0) {
                    setInfoMark(mapInfoMarks, oUserTaskOcc, getString("WRG_USER_TASK"), Constants.MODEL_INFO_WARNING);
                }
            }
        } else if (numberOfOrgElements > 1) {
            setInfoMark(mapInfoMarks, oUserTaskOcc, getString("WRG_USER_TASK"), Constants.MODEL_INFO_WARNING);
        }
    }
}

function check_ServiceTask(mapInfoMarks, oModel) {
    var oServiceTaskOccs = getServiceTasksInModel(oModel);
    for (var i = 0; i < oServiceTaskOccs.length; i++) {
        var oServiceTaskOcc = oServiceTaskOccs[i];

        // Service task is supported by software service type and software service operation type
        var oCxns2SSTs = getCxns2SSTs(oServiceTaskOcc);
        var oCxns2SSOTs = getCxns2SSOTs(oServiceTaskOcc);
        var oCxns2OrgElements = getCxns2OrgElements(oServiceTaskOcc);

        if (checkISService(oCxns2SSTs)) {
            setInfoMark(mapInfoMarks, oServiceTaskOcc, getString("ERR_IS_SERVICE_TASK"), Constants.MODEL_INFO_ERROR);
        } else {
            if (oCxns2SSTs.length == 1 && oCxns2SSOTs.length == 1) {
                if (!checkSSOT2SST_Relation(oCxns2SSOTs, oCxns2SSTs)) {
                    setInfoMark(mapInfoMarks, oServiceTaskOcc, getString("ERR_SST_SSOT_NOT_RELATED"), Constants.MODEL_INFO_ERROR);
                }
            }
            if (oCxns2SSTs.length == 0 ||
                oCxns2SSOTs.length == 0 ||
                oCxns2OrgElements.length == 1) {
                if (BPMN2XML === false)
                    setInfoMark(mapInfoMarks, oServiceTaskOcc, getString("WRG_SERVICE_TASK"), Constants.MODEL_INFO_WARNING);
            }
        }
    }
}


function checkISService(oCxns2SSTs){
    if (oCxns2SSTs.length==0) return false;
    var oSST = oCxns2SSTs[0].SourceObjDef();
    var oAttr =  oSST.Attribute(Constants.AT_SERVICE_TYPE, gn_Lang);
    if (oAttr!=null && oAttr.MeasureUnitTypeNum() == Constants.AVT_IS_SERVICE) return true;
    return false;
}

function checkSSOT2SST_Relation(oCxns2SSOTs,oCxns2SSTs){
    var oSSOT = oCxns2SSOTs[0].SourceObjDef();
    var oSST = oCxns2SSTs[0].SourceObjDef();
    var oSSTx = oSSOT.getConnectedObjs([Constants.OT_APPL_SYS_TYPE],Constants.EDGES_IN,[Constants.CT_CAN_SUBS_2]);
    return checkSSOT2SST_Relation2(oSSOT,oSST, []);
}

function checkSSOT2SST_Relation2(p_oSSOT,p_oSST, p_aVisited){
    var oSSTx = p_oSSOT.getConnectedObjs([Constants.OT_APPL_SYS_TYPE],Constants.EDGES_IN,[Constants.CT_CAN_SUBS_2]);
    for (var i = 0; i < oSSTx.length; i++){
        var oObj = oSSTx[i];
        if (oObj.equals(p_oSST)) return true;
    }
    for (var i = 0; i < oSSTx.length; i++){
        var oObj = oSSTx[i];
        if (p_aVisited.indexOf(oObj)>=0) continue;
        p_aVisited.push(oObj);
        if (checkSSOT2SST_Relation2(oObj,p_oSST,p_aVisited)) return true
    }
    return false
}

function check_AbstractTask(mapInfoMarks, oModel) {
    var oAbstractTaskOccs = getAbstractTasksInModel(oModel);
    for (var i = 0; i < oAbstractTaskOccs.length; i++) {
        var oAbstractTaskOcc = oAbstractTaskOccs[i];

        // Abstract task should be replaced by symbol manual task, user task or service task
        setInfoMark(mapInfoMarks, oAbstractTaskOcc, getString("WRG_ABSTRACT_TASK"), Constants.MODEL_INFO_WARNING);
    }
}

function check_SpecialTaskRelations(mapInfoMarks, oModel) {
    var oCallActivityOccs = getCallActivitiesInModel(oModel);         // Call Activities
    var oTaskOccs = oCallActivityOccs.concat(getSubprocessesInModel(oModel));   // Subprocesses

    for (var i = 0; i < oTaskOccs.length; i++) {
        var oTaskOcc = oTaskOccs[i];

        if (!TaskHasNoRelations(oTaskOcc)) {
            setInfoMark(mapInfoMarks, oTaskOcc, getString("ERR_WRONG_TASK_RELATIONS"), Constants.MODEL_INFO_ERROR);
        }
    }
}

function TaskHasNoRelations(oTaskOcc) {
    var oCxns2OrgElements   = getCxns2OrgElements(oTaskOcc);
    var oCxns2SSTs          = getCxns2SSTs(oTaskOcc);
    var oCxns2SSOTs         = getCxns2SSOTs(oTaskOcc);

    return (oCxns2OrgElements.length == 0 &&
        oCxns2SSTs.length == 0 &&
        oCxns2SSOTs.length == 0);
}

function check_SyncAgainstCentraSite(mapInfoMarks, oModel) {
    // Synchronisation against CentraSite
    var wMComp = Context.getComponent("webMethodsIntegration");

    var oServiceTaskOccs = getServiceTasksInModel(oModel);

    for (var i = 0; i < oServiceTaskOccs.length; i++) {
        var oTaskOcc = oServiceTaskOccs[i];

        var oServiceDefs = getConnectedServices(oTaskOcc);
        for (var j = 0; j < oServiceDefs.length; j++) {
            var oServiceDef = oServiceDefs[j];
            var result = wMComp.isManagedByCentraSite(oServiceDef);
            if (result == 0) {
                // Service is not synchronized against CentraSite (= 0)
                var oServiceOccs = oServiceDef.OccList();
                for (var k = 0; k < oServiceOccs.length; k++) {
                    setInfoMark( mapInfoMarks, oServiceOccs[k], getString("ERR_SYNC_AGAINST_CS"), Constants.MODEL_INFO_ERROR);
                }
            }
            if (result == 2) {
                // Service is currently in the worklist (= 2)
                var oServiceOccs = oServiceDef.OccList();
                for (var k = 0; k < oServiceOccs.length; k++) {
                    setInfoMark( mapInfoMarks, oServiceOccs[k], getString("ERR_SYNC_AGAINST_CS_2"), Constants.MODEL_INFO_ERROR);
                }
            }
        }
    }

    function getConnectedServices(oTaskOcc) {
        var oServiceDefs = new Array();

        var oCxns2SSTs = getCxns2SSTs(oTaskOcc);
        for (var i = 0; i < oCxns2SSTs.length; i++) {
            oServiceDefs.push(getConnectedService(oCxns2SSTs[i], oTaskOcc));
        }

        var oCxns2SSOTs = getCxns2SSOTs(oTaskOcc);
        for (var i = 0; i < oCxns2SSOTs.length; i++) {
            oServiceDefs.push(getConnectedService(oCxns2SSOTs[i], oTaskOcc));
        }
        return oServiceDefs;

        function getConnectedService(oCxn, oTaskOcc) {
            if (oTaskOcc.ObjDef().IsEqual(oCxn.SourceObjDef())) {
                return oCxn.TargetObjDef();
            }
            return oCxn.SourceObjDef();
        }
    }
}

function check_Uniqueness(mapInfoMarks, aModels) {
    var aRefOT = [Constants.OT_BPMN_POOL];

    /**
     * Determine whether the given object occurence is a call activity.
     *
     * @param objOcc
     * @returns {boolean} true, if attribute AT_BPMN_ACTIVITY_TYPE is maintained and set to AVT_BPMN_CALL_ACTIVITY.
     */
    function isCallActivity(objOcc) {
        var objDef = objOcc.ObjDef();
        var typeNum = objDef.TypeNum();
        if (typeNum == Constants.OT_FUNC) {
            var activityTypeAttr = objDef.Attribute(Constants.AT_BPMN_ACTIVITY_TYPE, gn_Lang);
            return activityTypeAttr.IsMaintained() && activityTypeAttr.MeasureUnitTypeNum() == Constants.AVT_BPMN_CALL_ACTIVITY;
        }
        return false;
    }

    for (var i = 0; i < aModels.length; i++) {
        var oSelModel = aModels[i];

        var diagramSets = [];
        getBPMNDiagramModelSet(oSelModel, diagramSets);
        
        var diagramList = [];
        for (var j = 0; j < diagramSets.length; j++) {
            var _diagrams = diagramSets[j];
            diagramList = diagramList.concat(_diagrams);
        }

        for (var j = 0; j < diagramSets.length; j++) {
            var diagramSet = diagramSets[j];

            for (var k = 0; k < diagramSet.length; k++) {
                var model = diagramSet[k];
                var objOccs = model.ObjOccList();

                // check whether occurrences in same diagram are unique or allowed.
                for (var m = 0; m < objOccs.length; m++) {
                    var objOcc = objOccs[m];
                    var typeNum = objOcc.ObjDef().TypeNum();

                    if (aRefOT.contains(typeNum)) {
                        var furtherOccs = getFurtherOccs(objOcc, diagramList);

                        if (furtherOccs.length > 0 && !isCallActivity(objOcc)) {
                            var message = getString("ERR_NOT_UNIQUE_OBJECT") + getAdditionalText(furtherOccs);
                            setInfoMark(mapInfoMarks, objOcc, message, Constants.MODEL_INFO_ERROR);
                        }
                    }
                }

                // check whether assignments of process models are unique
                if (model.OrgModelTypeNum() != Constants.MT_BPMN_PROCESS_DIAGRAM) continue;
                var aOccsWithAssignment = getOccsWithAssignment(model, diagramSet);
                for (var n = 0; n < aOccsWithAssignment.length; n++) {
                    var objOcc = aOccsWithAssignment[n];
                    var sText = formatstring1(getString("ERR_NOT_UNIQUE_ASSIGNMENT"), model.Name(gn_Lang)) + getAdditionalText2(objOcc, aOccsWithAssignment);
                    setInfoMark(mapInfoMarks, objOcc, sText, Constants.MODEL_INFO_ERROR);
                }
            }
        }
    }

    function getBPMNDiagramModelSet(model, modelSets, index) {
        if (typeof index == 'undefined') {
            for (var i = 0; i < modelSets.length; i++) {
                if (modelSets[i].length > 0) {
                    var mainModel = modelSets[i][0];
                    if (mainModel.IsEqual(model)) return;
                }
            }
            index = modelSets.length;
            modelSets.push([model]);
        } else if (modelSets[index].contains(model)) {
            return;
        } else {
            modelSets[index].push(model);
        }

        var functionDefs = model.ObjDefListFilter(Constants.OT_FUNC);
        for (var i = 0; i < functionDefs.length; i++) {
            var functionDef = functionDefs[i];

            if (isCallActivityWithGlobalProcess(functionDef)) {
                var assignedModels = getAssignedModelsOfBasType(functionDef, [Constants.MT_BPMN_COLLABORATION_DIAGRAM]);
                if (assignedModels.length > 0) {
                    getBPMNDiagramModelSet(assignedModels[0], modelSets);
                }
            } else {
                var assignedModels = getAssignedModelsOfBasType(functionDef, [Constants.MT_BPMN_COLLABORATION_DIAGRAM, Constants.MT_BPMN_PROCESS_DIAGRAM]);
                for (var j = 0; j < assignedModels.length; j++) {
                    getBPMNDiagramModelSet(assignedModels[j], modelSets, index);
                }
            }
        }
    }

    function isCallActivityWithGlobalProcess(objDef) {
        var oAttr = objDef.Attribute(Constants.AT_BPMN_CALLED_ELEMENT, gn_Lang);
        return (oAttr.IsMaintained() && oAttr.MeasureUnitTypeNum() == Constants.AVT_BPMN_GLOBAL_PROCESS)
    }

    function getFurtherOccs(oObjOcc, aModelList) {
        var oFurtherOccs = [];

        for (var i = 0; i < aModelList.length; i++) {
            var oModel = aModelList[i];

            var oOccList = oObjOcc.ObjDef().OccListInModel(oModel);
            for (var j = 0; j < oOccList.length; j++) {
                var oCurrOcc = oOccList[j];
                if (!oCurrOcc.IsEqual(oObjOcc)) oFurtherOccs.push(oCurrOcc);
            }
        }
        return oFurtherOccs;
    }

    function getOccsWithAssignment(p_oModel, aModelList) {
        var bNumberOfAssignments = 0;
        var oOccsWithAssignment = [];

        var superiorObjDefs = p_oModel.SuperiorObjDefs();
        if (superiorObjDefs.length > 1) {

            for (var i = 0; i < superiorObjDefs.length; i++) {
                var superiorObjDef = superiorObjDefs[i];

                for (var j = 0; j < aModelList.length; j++) {
                    var oModel = aModelList[j];
                    var oFuncDefList = oModel.ObjDefListFilter(Constants.OT_FUNC);

                    if (isItemInList(superiorObjDef, oFuncDefList)) {
                        bNumberOfAssignments++;
                        oOccsWithAssignment = oOccsWithAssignment.concat(superiorObjDef.OccListInModel(oModel));
                    }
                }
            }
        }
        return oOccsWithAssignment;
    }

    function getAdditionalText(oFurtherOccs) {
        var oModels = new Array();
        for (var i = 0; i < oFurtherOccs.length; i++) {
            oModels.push(oFurtherOccs[i].Model())
        }
        oModels = ArisData.Unique(oModels);
        oModels = ArisData.sort(oModels, Constants.AT_NAME, gn_Lang);

        var sText = "<ul>";
        for (var i = 0; i < oModels.length; i++) {
            sText += "<li>" + oModels[i].Name(gn_Lang) + "</li>";
        }
        sText += "</ul>";
        return sText;
    }

    function getAdditionalText2(oCurrOcc, oOccsWithAss) {
        var oObjDefs = new Array();
        for (var i = 0; i < oOccsWithAss.length; i++) {
            var oObjDefWithAss = oOccsWithAss[i].ObjDef()
            if (!oCurrOcc.ObjDef().IsEqual(oObjDefWithAss)) {
                oObjDefs.push(oObjDefWithAss)
            }
        }
        oObjDefs = ArisData.Unique(oObjDefs);
        oObjDefs = ArisData.sort(oObjDefs, Constants.AT_NAME, gn_Lang);

        var sText = "<ul>";
        for (var i = 0; i < oObjDefs.length; i++) {
            sText += "<li>" + oObjDefs[i].Name(gn_Lang) + "</li>";
        }
        sText += "</ul>";
        return sText;
    }
}

function check_AssignedFadModel(mapInfoMarks, oModel) {
    var oFuncOccs = oModel.ObjOccListFilter(Constants.OT_FUNC);
    for (var i = 0; i < oFuncOccs.length; i++) {
        var oFuncOcc = oFuncOccs[i];
        var oAssignedFadModels = getAssignedModelsOfBasType(oFuncOcc.ObjDef(), [Constants.MT_FUNC_ALLOC_DGM]);
        for (var j = 0; j < oAssignedFadModels.length; j++) {

            var oFadModel = oAssignedFadModels[j];
            var oAssFuncOccs = oFadModel.ObjOccListFilter(Constants.OT_FUNC);
            for (var k = 0; k < oAssFuncOccs.length; k++) {
                var oAssFuncOcc = oAssFuncOccs[k];

                if (!oAssFuncOcc.ObjDef().IsEqual(oFuncOcc.ObjDef())) {
                    setInfoMark(mapInfoMarks, oAssFuncOcc, getString("ERR_ASSIGNED_FAD"), Constants.MODEL_INFO_ERROR);
                }
            }
        }
    }
}

function checkTaskHierarchyCycle(mapInfoMarks, oModel) {
    var oTaskOccs = oModel.ObjOccListFilter(Constants.OT_FUNC);
    var bCycle = false;
    for (var i = 0; i < oTaskOccs.length; i++) {
        var oTaskOcc = oTaskOccs[i];
        var oHierarchyModels = new Array();

        if (isCycleInHierarchy(oTaskOcc, oHierarchyModels)) {
            setInfoMark(mapInfoMarks, oTaskOcc, getString("ERR_HIERARCHY_CYCLE"), Constants.MODEL_INFO_ERROR);
            bCycle = true;
        }
    }
    return bCycle;

    function isCycleInHierarchy(oObjOcc, oHierarchyModels) {
        var oAssignedModels = getAssignedModelsOfBasType(oObjOcc.ObjDef(), [Constants.MT_BPMN_COLLABORATION_DIAGRAM, Constants.MT_BPMN_PROCESS_DIAGRAM]);
        if (oAssignedModels.length > 0) {
            for (var j = 0; j < oAssignedModels.length; j++) {
                var oModel = oAssignedModels[j];
                if (isItemInList(oModel, oHierarchyModels)) return true;

                var oNewHierarchyModels = getCopiedModelList(oHierarchyModels);
                oNewHierarchyModels.push(oModel);

                var oTaskOccs = oModel.ObjOccListFilter(Constants.OT_FUNC);
                for (var k = 0; k < oTaskOccs.length; k++) {
                    var oTaskOcc = oTaskOccs[k];

                    if (isCycleInHierarchy(oTaskOcc, oNewHierarchyModels)) return true;
                }
            }
        }
        return false;

        function getCopiedModelList(oModelList) {
            var oCopiedModelList = new Array();
            for (var i = 0; i < oModelList.length; i++) {
                oCopiedModelList.push(oModelList[i]);
            }
            return oCopiedModelList;
        }
    }

}

function checkTaskAssignment(mapInfoMarks, oModel) {
    var oCallActivityOccs = getCallActivitiesInModel(oModel);
    var oSubprocessOccs = getSubprocessesInModel(oModel);

    var oObjOccs = oModel.ObjOccListFilter(Constants.OT_FUNC);
    for (var i = 0; i < oObjOccs.length; i++) {
        var oObjOcc = oObjOccs[i];
        var wrongAssignments = new Array();

        // Assignment with 'Collaboration Diagram' or 'Process Diagram' - only allowed for call activities and subprocesses
        var oAssignedModels = getAssignedModelsOfBasType(oObjOcc.ObjDef(), [Constants.MT_BPMN_COLLABORATION_DIAGRAM, Constants.MT_BPMN_PROCESS_DIAGRAM]);
        if (oAssignedModels.length > 0) {
            for (var j = 0; j < oAssignedModels.length; j++) {
                var oAssignedModel = oAssignedModels[j];
                if (oAssignedModel.OrgModelTypeNum() == Constants.MT_BPMN_COLLABORATION_DIAGRAM) {
                    if (!isOccInList(oObjOcc, oCallActivityOccs)) wrongAssignments.push(oAssignedModel);
                }
                else if (oAssignedModel.OrgModelTypeNum() == Constants.MT_BPMN_PROCESS_DIAGRAM) {
                    if (!isOccInList(oObjOcc, oSubprocessOccs)) wrongAssignments.push(oAssignedModel);
                }
            }
        }
        if (wrongAssignments.length > 0) {
            var sText = getString("ERR_TASK_ASSIGNMENT") + getAdditionalText(wrongAssignments);
            setInfoMark(mapInfoMarks, oObjOcc, sText, Constants.MODEL_INFO_ERROR);
        }
    }
    // Assignment with 'FAD' - NOT allowed for call activities and subprocesses
    oObjOccs = oCallActivityOccs.concat(oSubprocessOccs);
    for (var i = 0; i < oObjOccs.length; i++) {
        var oObjOcc  = oObjOccs[i];
        wrongAssignments = new Array();
        var oAssignedModels = getAssignedModelsOfBasType(oObjOcc.ObjDef(), [Constants.MT_FUNC_ALLOC_DGM]);
        if (oAssignedModels.length > 0) {
            for (var j = 0; j < oAssignedModels.length; j++) {
                var oAssignedModel = oAssignedModels[j];
                wrongAssignments.push(oAssignedModel);
            }
        }
        if (wrongAssignments.length > 0) {
            var sText = getString("ERR_TASK_ASSIGNMENT") + getAdditionalText(wrongAssignments);
            setInfoMark(mapInfoMarks, oObjOcc, sText, Constants.MODEL_INFO_ERROR);
        }
    }

    function isOccInList(oObjOcc, oOccList) {
        for (var i = 0; i < oOccList.length; i++) {
            if (oOccList[i].IsEqual(oObjOcc)) return true;
        }
        return false;
    }

    function getAdditionalText(oModels) {
        oModels = ArisData.Unique(oModels);
        oModels = ArisData.sort(oModels, Constants.AT_NAME, gn_Lang);

        var sText = "<ul>";
        for (var i = 0; i < oModels.length; i++) {
            sText += "<li>" + formatstring2(getString("NAME_TYPE"), oModels[i].Name(gn_Lang), oModels[i].Type()) + "</li>";
        }
        sText += "</ul>";
        return sText;
    }
}

function check_wMChecks(mapInfoMarks, oModel) {
    // This function runs the wM semantic checks, implemented in java class "ABPMNSemanticCheck"

    if (BPMN2XML == null || BPMN2XML === false)
        runCheck("wMCheck_GatewayEventBased", mapInfoMarks, oModel, getString("ERR_EVENTBASED_GATEWAY_NOT_SUPPORTED"), Constants.MODEL_INFO_ERROR);

    // Not supported start events
    runCheck_NotSupportedEvents("wMCheck_StartEventMultiple_TopLevel");
    runCheck_NotSupportedEvents("wMCheck_StartEventParallelMultiple_TopLevel");
    runCheck_NotSupportedEvents("wMCheck_StartEventMultiple_EventSubProcessI");
    runCheck_NotSupportedEvents("wMCheck_StartEventParallelMultiple_EventSubProcessI");
    runCheck_NotSupportedEvents("wMCheck_StartEventMultiple_EventSubProcessNI");
    runCheck_NotSupportedEvents("wMCheck_StartEventParallelMultiple_EventSubProcessNI");
    // Not supported intermediate events    
    runCheck_NotSupportedEvents("wMCheck_IntermediateEventMultiple_Catching");
    runCheck_NotSupportedEvents("wMCheck_IntermediateEventParallelMultiple_Catching");
    runCheck_NotSupportedEvents("wMCheck_IntermediateEventMultiple_BoundaryI");
    runCheck_NotSupportedEvents("wMCheck_IntermediateEventParallelMultiple_BoundaryI");
    runCheck_NotSupportedEvents("wMCheck_IntermediateEventMultiple_BoundaryNI");
    runCheck_NotSupportedEvents("wMCheck_IntermediateEventParallelMultiple_BoundaryNI");
    runCheck_NotSupportedEvents("wMCheck_IntermediateEventMultiple_Throwing");
    // Not supported end events
    runCheck_NotSupportedEvents("wMCheck_EndEventMultiple_TopLevel");

    function runCheck_NotSupportedEvents(checkName) {
        runCheck(checkName, mapInfoMarks, oModel, getString("ERR_NOT_SUPPORTED_EVENTS"), Constants.MODEL_INFO_ERROR);
    }

    function runCheck(checkName, mapInfoMarks, oModel, sText, nType) {
        var result = g_semCheckComponent.runCheck(checkName, oModel);
        handleResult(oModel, result, sText);

        function handleResult(oModel, result, sText) {
            if (result == null || !result.wasChecked()) {
                return;
            }
            var errors = result.getErrors();
            for (var i = 0; i < errors.length; i++) {
                var error = errors[i];
                var oObjOcc = error.getOcc();

                setInfoMark(mapInfoMarks, oObjOcc, sText, nType);
            }
        }
    }
}

function getCxns2OrgElements(oTaskOcc){
    var oCxns2OrgElements = new Array();

    var oAssignedTaskOcc = getOccInAssignedModel(oTaskOcc.ObjDef());
    if (oAssignedTaskOcc == null) return oCxns2OrgElements;

    oCxns2OrgElements = oCxns2OrgElements.concat(getRelevantCxns(oAssignedTaskOcc, Constants.CT_EXEC_1, [Constants.ST_POS], Constants.EDGES_IN));
    oCxns2OrgElements = oCxns2OrgElements.concat(getRelevantCxns(oAssignedTaskOcc, Constants.CT_EXEC_2, [Constants.ST_EMPL_TYPE], Constants.EDGES_IN));
    oCxns2OrgElements = oCxns2OrgElements.concat(getRelevantCxns(oAssignedTaskOcc, Constants.CT_EXEC_1, [Constants.ST_ORG_UNIT_1, Constants.ST_ORG_UNIT_2], Constants.EDGES_IN));
    oCxns2OrgElements = oCxns2OrgElements.concat(getRelevantCxns(oAssignedTaskOcc, Constants.CT_EXEC_1, [Constants.ST_GRP], Constants.EDGES_IN));
    oCxns2OrgElements = oCxns2OrgElements.concat(getRelevantCxns(oAssignedTaskOcc, Constants.CT_EXEC_2, [Constants.ST_ORG_UNIT_TYPE_1, Constants.ST_ORG_UNIT_TYPE_2], Constants.EDGES_IN));
    return ArisData.Unique(oCxns2OrgElements);
}

function getCxns2SSTs(oTaskOcc){
    var oCxns2SSTs = new Array();

    var oAssignedTaskOcc = getOccInAssignedModel(oTaskOcc.ObjDef());
    if (oAssignedTaskOcc == null) return oCxns2SSTs;

    oCxns2SSTs = oCxns2SSTs.concat(getRelevantCxns(oAssignedTaskOcc, Constants.CT_CAN_SUPP_1, [Constants.ST_SW_SERVICE_TYPE], Constants.EDGES_IN));
    return ArisData.Unique(oCxns2SSTs);
}

function getCxns2SSOTs(oTaskOcc){
    var oCxns2SSOTs = new Array();

    var oAssignedTaskOcc = getOccInAssignedModel(oTaskOcc.ObjDef());
    if (oAssignedTaskOcc == null) return oCxns2SSOTs;

    oCxns2SSOTs = oCxns2SSOTs.concat(getRelevantCxns(oAssignedTaskOcc, Constants.CT_SUPP_3, [Constants.ST_SW_SERVICE_OPERATION_TYPE], Constants.EDGES_IN));
    return ArisData.Unique(oCxns2SSOTs);
}

function getOccInAssignedModel(oObjDef){
    var oAssignedModels = getAssignedModelsOfBasType(oObjDef, [Constants.MT_FUNC_ALLOC_DGM]);
    if (oAssignedModels.length > 0) {
        var oOccsInAssignedModel = oObjDef.OccListInModel(oAssignedModels[0]);
        if (oOccsInAssignedModel.length > 0) {

            return oOccsInAssignedModel[0];
        }
    }
    return null;
}

function getRelevantCxns(oUserTaskOcc, nCxnType, aConnSymbols, nInOut) {
    var oRelevantCxns = new Array();
    var oCxnOccs = oUserTaskOcc.Cxns(nInOut);
    for (var i = 0; i < oCxnOccs.length; i++) {
        var oCxnOcc = oCxnOccs[i];
        if (oCxnOcc.Cxn().TypeNum() == nCxnType) {
            var oConnObjOcc = (nInOut == Constants.EDGES_IN) ? oCxnOcc.SourceObjOcc() : oCxnOcc.TargetObjOcc();
            for (var j = 0; j < aConnSymbols.length; j++) {
                if (oConnObjOcc.OrgSymbolNum() == aConnSymbols[j]) {
                    oRelevantCxns.push(oCxnOcc.Cxn());
                }
            }
        }
    }
    return oRelevantCxns;
}

function getSelectedBPMNModels(){
    return getFilteredModels(ArisData.getSelectedModels());

    // Filters the models with the specific type nums
    function getFilteredModels(selectedModels) {
        var aTypeNums = Context.getDefinedItemTypes(Constants.CID_MODEL);
        if (aTypeNums.length == 0 || (aTypeNums.length == 1 && aTypeNums[0] == -1)) {
            // All/None type nums selected
            return selectedModels;
        }
        var setTypeNums = new java.util.HashSet();
        for (var i = 0; i < aTypeNums.length; i++) {
            setTypeNums.add(java.lang.Integer.valueOf(aTypeNums[i]));
        }
        var filteredModels = new Array();
        for (var i = 0; i < selectedModels.length; i++) {
            var oModel = selectedModels[i];
            if (setTypeNums.contains(java.lang.Integer.valueOf(oModel.OrgModelTypeNum()))) {
                filteredModels.push(oModel);
            }
        }
        return filteredModels;
    }
}

function initInfoMarks(oModels) {
    var mapInfoMarks = new java.util.HashMap();
    for (var i = 0; i < oModels.length; i++) {
        var key1 = oModels[i].GUID();
        if (!mapInfoMarks.containsKey(key1)) {
            mapInfoMarks.put(key1, new java.util.HashMap());
        }
    }
    return mapInfoMarks;
}

function setInfoMark(mapInfoMarks, oOcc, sText, nType) {

    sText = "<li>" + sText + "</li>";

    var sErrors = "";
    var sWarnings = "";
    var sInfos = ""
    if(nType == Constants.MODEL_INFO_ERROR) {
        gb_c_ERROR_MARK_SET = true;
        sErrors = sText;
    }
    if(nType == Constants.MODEL_INFO_WARNING) {
        gb_c_WARNING_MARK_SET = true;
        sWarnings = sText;
    }
    if(nType == Constants.MODEL_INFO_INFORMATION) {
        gb_c_INFO_MARK_SET = true;
        sInfos = sText;
    }

    var key1 = oOcc.Model().GUID();
    var key2 = oOcc.ObjectID();

    var value1 = new java.util.HashMap(); // = inner HashMap

    if (mapInfoMarks.containsKey(key1)) {
        value1 = mapInfoMarks.get(key1);
        if (value1.containsKey(key2)) {
            var value2 = value1.get(key2);

            sErrors = value2.sErrors + sErrors;
            sWarnings = value2.sWarnings + sWarnings;
            sInfos = value2.sInfos + sInfos;

            nType = Math.max(nType, value2.nType);
        }
    }
    var value2 = new INFO(sErrors, sWarnings, sInfos, nType);

    value1.put(key2, value2);
    mapInfoMarks.put(key1, value1);
}

function outInfoMarks(mapInfoMarks) {
    if( gb_DGRMS_OPEN_SET == true ) {
        var modelSet = mapInfoMarks.keySet();
        var iter = modelSet.iterator();
        while (iter.hasNext()) {
            var key1 = iter.next();
            var value1 = mapInfoMarks.get(key1);

            var oModel = g_oDB.FindGUID(key1, Constants.CID_MODEL);
            if (value1.keySet().size() > 0) {
                oModel.openModel();
            }
            oModel.ClearError();

            var occSet = value1.keySet();
            var iter2 = occSet.iterator();
            while (iter2.hasNext()) {
                var key2 = iter2.next();
                var value2 = value1.get(key2);

                var sErrors = ""+value2.sErrors;
                var sWarnings = ""+value2.sWarnings;
                var sInfos = ""+value2.sInfos;

                var sText = "";
                if (sErrors.length > 0) {
                    sText += "<b>" + getString("ERRORS") + "</b><ul>" + sErrors + "</ul>";
                }
                if (sWarnings.length > 0) {
                    sText += "<b>" + getString("WARNINGS") + "</b><ul>" + sWarnings + "</ul>";
                }
                if (sInfos.length > 0) {
                    sText += "<b>" + getString("INFOS") + "</b><ul>" + sInfos + "</ul>";
                }
                sText = "<html>" + sText.replace(/\n/g, "<br>") + "</html>";
                var nType = value2.nType;

                var oOcc = g_oDB.FindOID(key2);
                oModel.setTemporaryUserInfo(oOcc, nType, sText);
            }
        }
    } else {
        // Clear errors if models are open
        var modelSet = mapInfoMarks.keySet();
        var iter = modelSet.iterator();
        while (iter.hasNext()) {
            var key1 = iter.next();
            var oModel = g_oDB.FindGUID(key1, Constants.CID_MODEL);
            oModel.ClearError();
        }
    }
}

function writeProperties( ){
    Context.setProperty("INFO_MARK_SET", gb_c_INFO_MARK_SET);
    Context.setProperty("WARNING_MARK_SET", gb_c_WARNING_MARK_SET);
    Context.setProperty("ERROR_MARK_SET", gb_c_ERROR_MARK_SET);
    Context.setProperty("DIAGRAMS_TO_BE_OPENED", gb_DGRMS_OPEN_SET);
}

function getBoolPropertyValue(p_sPropKey) {
    var property = Context.getProperty(p_sPropKey);
    if (property != null) {
        return (StrComp(property, "true") == 0);
    }
    return false;
}

function isItemInList(oItem, oItemList) {
    for (var i = 0; i < oItemList.length; i++) {
        if (oItemList[i].IsEqual(oItem)) return true;
    }
    return false;
}
/*********************************************************************************************/

function getSubprocessesInModel(p_oModel) {
    return getOccsInModel_wMHelper("wMHelper_getSubprocesses", p_oModel);
}

function getCallActivitiesInModel(p_oModel) {
    return getOccsInModel_wMHelper("wMHelper_getCallActivities", p_oModel);
}

function getManualTasksInModel(p_oModel) {
    return getOccsInModel_wMHelper("wMHelper_getManualTasks", p_oModel);
}

function getUserTasksInModel(p_oModel) {
    return getOccsInModel_wMHelper("wMHelper_getUserTasks", p_oModel);
}

function getServiceTasksInModel(p_oModel) {
    return getOccsInModel_wMHelper("wMHelper_getServiceTasks", p_oModel);
}

function getAbstractTasksInModel(p_oModel) {
    return getOccsInModel_wMHelper("wMHelper_getAbstractTasks", p_oModel);
}

function getOccsInModel_wMHelper(checkName, oModel) {
    var result = g_semCheckComponent.runCheck(checkName, oModel);
    if (result == null || !result.wasChecked()) {
        return new Array();
    }
    var oOccList = new Array();

    var errors = result.getErrors();
    for (var i = 0; i < errors.length; i++) {
        var oObjOcc = errors[i].getOcc();
        oOccList.push(oObjOcc);
    }
    return oOccList;
}

function checkOpenMergeConflicts(){
    var bModsOpened = false;
    var aModsChecked = [];

    this.getAllMods2Open = function(p_aMods2Check){
        if (aModsChecked.length == 0) aModsChecked = aModsChecked.concat(p_aMods2Check);
        for (var i=0; i<p_aMods2Check.length; i++){
            var oMod2Check = p_aMods2Check[i];
            if (mergeNotComplete(oMod2Check)){
                oMod2Check.openModel();
                bModsOpened = true;
            }
            this.getAllMods2Open(getAllUnCheckAssigments(oMod2Check));
        }
        return bModsOpened;
    }
    function mergeNotComplete(p_oMod2Check){
        var oMergeState = p_oMod2Check.Attribute(Constants.AT_MOD_TRANS_RESOLVE_STATE ,-1);
        var bMergedNotCompleted = (oMergeState.IsMaintained()) ? oMergeState.MeasureUnitTypeNum ( ) != Constants.AVT_MOD_TRANSFORM_RESOLVE_COMPLETED : false;
        return  bMergedNotCompleted;
    }
    function getAllUnCheckAssigments(p_oMod2Check){
        var aMods2Check = [];
        var aFkts = p_oMod2Check.ObjDefListFilter(Constants.OT_FUNC);
        for (var i=0 ; i < aFkts.length; i++){
            var aAssignedMods = getAssignedModelsOfBasType(aFkts[i], [Constants.MT_BPMN_COLLABORATION_DIAGRAM, Constants.MT_BPMN_PROCESS_DIAGRAM, Constants.MT_FUNC_ALLOC_DGM]);
            for (var j=0; j < aAssignedMods.length; j++){
                if (aAssignedMods[j].OrgModelTypeNum() == Constants.MT_FUNC_ALLOC_DGM){
                    if (mergeNotComplete(aAssignedMods[j])){
                        bModsOpened = true;
                        aAssignedMods[j].openModel();
                    }
                } else {
                    if (!aModsChecked.contains(aAssignedMods[j])){
                        aMods2Check.push(aAssignedMods[j]);
                    }
                    aModsChecked.push(aAssignedMods[j]);
                }
            }
        }
        return aMods2Check;
    }
}

function getAssignedModelsOfBasType(object, baseTypes) {
    var models = object.AssignedModels(null);
    var findings = [];

    for (var j = 0; j < baseTypes.length; j++) {
        var baseType = baseTypes[j];
        if (typeof baseType == 'undefined') {
            throw Error("baseTypes[" + j + "] is undefined!");
        }
    }

    for (var i = 0; i < models.length; i++) {
        var model = models[i];
        if (baseTypes.contains(model.OrgModelTypeNum())) {
            findings.push(model);
        }
    }

    return findings;
}
