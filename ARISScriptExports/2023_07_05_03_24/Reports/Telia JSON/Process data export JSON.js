var nLocale = Context.getSelectedLanguage()
var outputFile = Context.createOutputObject()

var selectedModels = ArisData.getSelectedModels()
var selectedModel = selectedModels[0]

var arrayJSON = new org.json.JSONArray();

var startParent = "000"
var variantParent = "000"
var modelName = selectedModel.Name(nLocale)
var modelGuid = selectedModel.GUID()
var hierarchyLevel = "L0" //first model always L0
var status = selectedModel.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("d2ac7740-af65-11df-599b-005056aa0ab6"), nLocale).getValue()
var typeName = selectedModel.Type()
var typeKind = "MODEL"
var href = getHref(selectedModel, "models")
var description = selectedModel.Attribute(Constants.AT_DESC, nLocale).getValue()

var modelJsonObject = createModelJsonObject(modelGuid, modelName, startParent, hierarchyLevel, status, typeName, typeKind, href, description)
arrayJSON.put(modelJsonObject)

var objOccListTemp = selectedModel.ObjOccListFilter(Constants.OT_FUNC)

var modelObjectsStart = new Array()
for (var i = 0; i < objOccListTemp.length; i++) {
    var objDefTemp = objOccListTemp[i].ObjDef()
    //check Hierarchy Level
    if (objDefTemp.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("49e3ac80-b16f-11e2-424c-005056aa0ab6"), nLocale).IsMaintained()) {
        modelObjectsStart.push(objOccListTemp[i])
    }
}

for (var ix = 0; ix < modelObjectsStart.length; ix++) {
    getDataHier(modelObjectsStart[ix], modelGuid, variantParent, 0)
}

//format json data
var splitlen = JSON.parse(arrayJSON);
var uniqueArray = [];
for (i = 0; i < splitlen.length; i++) {
    if (!uniqueArray.find(compare)) {
        uniqueArray.push(splitlen[i]);
    }
}

var uniqueJsonArray = JSON.stringify(uniqueArray, 0, 4)

var flatArray = JSON.parse(uniqueJsonArray.toString(2));
var nested = getNestedChildren(flatArray, startParent)

outputFile.OutputTxt(JSON.stringify(nested[0], 0, 4));
outputFile.WriteReport();

function getDataHier(objOcc, parent, variantParent, variantDepth) {
    var objDef = objOcc.ObjDef()

    var id = objDef.GUID()
    var o_name = objDef.Name(nLocale)
    var description = objDef.Attribute(Constants.AT_DESC, nLocale).getValue()
    var hierarchyLevel = objDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("49e3ac80-b16f-11e2-424c-005056aa0ab6"), nLocale).getValue()
    var href = getHref(objDef, "objects")
    var variantAttribute = objDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("f1b07770-7855-11ec-7d84-00505682bf5e"), nLocale).getValue()

    var officialStatusAtt = objDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("768c6660-45c9-11df-599b-005056aa0ab6"), nLocale)
    if (officialStatusAtt.IsMaintained()) {
        var officialStatus = objDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("768c6660-45c9-11df-599b-005056aa0ab6"), nLocale).getValue()
    }
    else {
        var officialStatus = null
    }

    var typeKind = "MODELOBJECT"
    var typeName = objDef.Type()

    //model data
    var occModel = objOcc.Model()
    var modelID = occModel.GUID()
    var modelHref = getHref(occModel, "models")
    var modelName = occModel.Name(nLocale)
    var modelStatusAtt = occModel.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("d2ac7740-af65-11df-599b-005056aa0ab6"), nLocale)
    if (modelStatusAtt.IsMaintained()) {
        var modelStatus = occModel.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("d2ac7740-af65-11df-599b-005056aa0ab6"), nLocale).getValue()
    }
    else {
        var modelStatus = null
    }
    var modelTypeName = occModel.Type()
    var modelTypeNum = occModel.TypeNum()

    //create model data JSON
    var modelJsonObject = createModelJsonObject(modelID, modelName, null, null, modelStatus, modelTypeName, null, modelHref, null)

    var bpmnAssigned = null
    var assignedBPMNModels = objDef.AssignedModels([Constants.MT_BPMN_COLLABORATION_DIAGRAM, Constants.MT_EEPC, ArisData.ActiveFilter().UserDefinedModelTypeNum("9f0ac401-0e0f-11e6-7464-0050560116c9")])
    if (assignedBPMNModels.length != 0) {
        var bpmnAssigned = "true"
    }

    //create json object
    if (modelTypeNum == Constants.MT_FUNC_TREE || modelTypeNum == ArisData.ActiveFilter().UserDefinedModelTypeNum("2972efc1-c981-11e4-1c0c-0050560116c6")) {
        var functionJsonObject = createJsonObject(id, o_name, parent, description, hierarchyLevel, href, typeKind, typeName, bpmnAssigned, variantAttribute, officialStatus, variantDepth, modelJsonObject)
        arrayJSON.put(functionJsonObject); //this is JSON ARRAY
    }

    //if variant
    if (modelTypeNum == Constants.MT_PRCS_SLCT_MTX) {
        variantDepth = variantDepth + 1

        //create variant array
        var variantOfId = parent //if object is variant then parent is variantOf id
        var database = ArisData.getActiveDatabase()
        var variantOfObjDef = database.FindGUID(parent)
        var variantOfName = variantOfObjDef.Name(nLocale)
        var variantOfJsonObject = createVariantOfJsonObject(variantOfId, variantOfName)

        var variantJsonObject = createVariantJsonObject(id, o_name, variantParent, description, href, variantAttribute, typeKind, typeName, officialStatus, hierarchyLevel, bpmnAssigned, variantDepth, modelJsonObject, variantOfJsonObject)
        arrayJSON.put(variantJsonObject);
    }

    //get recursion objects
    var assignedModels = objDef.AssignedModels([Constants.MT_FUNC_TREE, Constants.MT_PRCS_SLCT_MTX])
    var funcObjOccs = new Array()
    for (var i = 0; i < assignedModels.length; i++) {
        var tempObjOccs = assignedModels[i].ObjOccListFilter(Constants.OT_FUNC)
        for (var j = 0; j < tempObjOccs.length; j++) {
            var tempObjDef = tempObjOccs[j].ObjDef()
            if (tempObjDef.GUID() == objDef.GUID()) {
                var connectedOccs = tempObjOccs[j].getConnectedObjOccs(0)
                for (k in connectedOccs) {
                    funcObjOccs.push(connectedOccs[k])
                }
            }
        }
    }

    //recursion
    for (var i = 0; i < funcObjOccs.length; i++) {
        getDataHier(funcObjOccs[i], id, parent, variantDepth)
    }
}

function createJsonObject(id, o_name, parent, description, hierarchyLevel, href, typeKind, typeName, bpmnAssigned, variantAttribute, officialStatus, variantDepth, objModel) {
    var objMain = new org.json.JSONObject();

    objMain.put("id", id);
    objMain.put("name", o_name);
    objMain.put("parent", parent);
    objMain.put("description", description);
    objMain.put("hierarchyLevel", hierarchyLevel);
    objMain.put("href", href);
    objMain.put("@typeKind", typeKind);
    objMain.put("@typeName", typeName);

    if (bpmnAssigned != null)
        objMain.put("bpmnAssigned", bpmnAssigned);

    if (officialStatus != null)
        objMain.put("status", officialStatus);

    if (variantAttribute == "Yes") {
        objMain.put("variant", variantAttribute);
        objMain.put("variantDepth", variantDepth);
    }

    objMain.put("model", objModel);

    return objMain;
}

function createModelJsonObject(m_id, m_name, m_parent, m_level, m_status, typeName, typeKind, href, description) {
    var obj = new org.json.JSONObject();

    obj.put("id", m_id);
    obj.put("name", m_name);

    if (m_status != null)
        obj.put("status", m_status);

    obj.put("@typeName", typeName);
    obj.put("href", href);

    if (m_level != null)
        obj.put("hierarchyLevel", m_level);
    if (m_parent != null)
        obj.put("parent", m_parent);
    if (typeKind != null)
        obj.put("@typeKind", typeKind);
    if (description != null)
        obj.put("description", description);

    return obj;
}

function createVariantOfJsonObject(v_id, v_name) {
    var obj = new org.json.JSONObject();

    obj.put("id", v_id);
    obj.put("name", v_name);

    return obj;
}

function createVariantJsonObject(v_id, v_name, parent, v_desc, v_href, v_att, v_typeKind, v_typeName, v_officialStatus, v_hierarchyLevel, v_bpmnAssigned, v_variantDepth, v_model, v_variantOf) {
    var objMain = new org.json.JSONObject();

    objMain.put("id", v_id);
    objMain.put("name", v_name);
    objMain.put("parent", parent);
    objMain.put("description", v_desc);

    objMain.put("href", v_href);

    if (v_att == "Yes")
        objMain.put("variant", v_att);

    objMain.put("variantDepth", v_variantDepth);
    objMain.put("@typeKind", v_typeKind);
    objMain.put("@typeName", v_typeName);

    if (v_officialStatus != null)
        objMain.put("status", v_officialStatus);

    if (v_bpmnAssigned != null)
        objMain.put("bpmnAssigned", v_bpmnAssigned);

    objMain.put("hierarchyLevel", v_hierarchyLevel);
    objMain.put("model", v_model);
    objMain.put("variantOf", v_variantOf);

    return objMain;
}

function getNestedChildren(arr, parent) {
    var out = []
    for (var i in arr) {
        if (arr[i].parent == parent) {
            var children = getNestedChildren(arr, arr[i].id)

            if (children.length) {
                arr[i].children = children
            }
            delete(arr[i].parent) //delete parent property
            out.push(arr[i])
        }
    }
    return out
}

function getHref(obj, type) {
    //https://aris.teliacompany.net/abs/api/objects/_TeliaCompany_WORKING/1e90941d-4922-11ec-0743-00505682361b?language=en_US&methodfilter=8f7b72e0-ffc4-11e0-50be-005056aa0ab6&methodlanguage=en_US
    var oDatabase = ArisData.getActiveDatabase();
    //var sTenantName = ArisData.getTenantName();
    var serverName = oDatabase.ServerName();
    var sDBName = oDatabase.Name(nLocale);
    sDBName = sDBName.replace(" ", "%20");
    var g_guid = obj.GUID()
    var filterGUID = oDatabase.ActiveFilter().FilterKey()

    var linkURL = "https://" + serverName + "/abs/api/" + type + "/" + sDBName + "/" + g_guid + "?language=en_US&methodfilter=" + filterGUID + "&methodlanguage=en_US"
    return linkURL
}

function compare(x) {
    if (x.id === splitlen[i].id && x.parent === splitlen[i].parent)
        return true
    else
        return false
}