var nLocale = Context.getSelectedLanguage()
var outputFile = Context.createOutputObject()

var selectedModels = ArisData.getSelectedModels()
var selectedModel = selectedModels[0]

var g_objectList = new Array();
var jsonString = ""
var arrayJSON = new org.json.JSONArray();

//GET START MODEL DATA
var zeroParent = "000"
var modelName = selectedModel.Name(nLocale)
var modelGuid = selectedModel.GUID()
var hierarchyLevel = "L0"
var status = selectedModel.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("d2ac7740-af65-11df-599b-005056aa0ab6"), nLocale).getValue()
var typeName = selectedModel.Type() //@typeName
var typeKind = "MODEL" //selectedModel.KindNum() //@typeKind
var href = getHref(selectedModel, "models") 
var description = selectedModel.Attribute(Constants.AT_DESC, nLocale).getValue()

var modelJsonObject = createModelJsonObject(modelGuid, modelName, zeroParent, hierarchyLevel, status, typeName, typeKind, href, description)
arrayJSON.put(modelJsonObject); //this is JSON ARRAY

//GET OBJECTS FROM THE START MODEL
//var objDefListTemp = selectedModel.ObjDefListFilter(Constants.OT_FUNC)
var objOccListTemp = selectedModel.ObjOccListFilter(Constants.OT_FUNC)

var modelObjectsStart = new Array()
for (var i = 0; i < objOccListTemp.length; i++) {
    var objDefTemp = objOccListTemp[i].ObjDef()
    //check Hierarchy Level att
    if (objDefTemp.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("49e3ac80-b16f-11e2-424c-005056aa0ab6"), nLocale).IsMaintained()) {
        modelObjectsStart.push(objOccListTemp[i])
    }
}

//go trought start objects and catch data
for (var ix = 0; ix < modelObjectsStart.length; ix++) {
    getDataHier(modelObjectsStart[ix], modelGuid)

}

//process json data
var flatArray = JSON.parse(arrayJSON.toString(2));
var nested = getNestedChildren(flatArray, zeroParent)

outputFile.OutputTxt(JSON.stringify(nested, 0, 4));
outputFile.WriteReport();

function getDataHier(objOcc, parent) {   
    var objDef = objOcc.ObjDef()
/*     if (existsInArray(objDef, g_objectList)) {
        return;
    } */
    g_objectList.push(objDef) //FOR CHECK, do not go twice in the same object

    var id = objDef.GUID()
    var o_name = objDef.Name(nLocale)
    var description =objDef.Attribute(Constants.AT_DESC, nLocale).getValue()
    var hierarchyLevel = objDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("49e3ac80-b16f-11e2-424c-005056aa0ab6"), nLocale).getValue()
    var href = getHref(objDef, "objects")
    var typeKind = "OBJECT"
    var typeName = objDef.Type()
    
    //OCCURENCE MODEL DATA
    var occModel = objOcc.Model()
    var modelID = occModel.GUID()
    var modelHref = getHref(occModel, "models")
    var modelName = occModel.Name(nLocale)
    var modelStatus = occModel.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("d2ac7740-af65-11df-599b-005056aa0ab6"), nLocale).getValue()
    var modelTypeName = occModel.Type()
    //create model data JSON
    var modelJsonObject = createModelJsonObject(modelID, modelName, null, null, modelStatus, modelTypeName, null, modelHref, null) 
    
    /* Set 'bpmnAssigned' to 'true' if assignment with 'apiname'/'typename' have:
    "MT_BPMN_COLLABORATION_DIAGRAM"/"BPMN collaboration diagram (BPMN 2.0)"
    "MT_EEPC"/"BPMN collaboration diagram (BPMN 2.0)"
    "9f0ac401-0e0f-11e6-7464-0050560116c9"/"Working process e2e 2.0 */
    var bpmnAssigned = "false"
    var assignedBPMNModels = objDef.AssignedModels([Constants.MT_BPMN_COLLABORATION_DIAGRAM, Constants.MT_EEPC, ArisData.ActiveFilter().UserDefinedModelTypeNum("9f0ac401-0e0f-11e6-7464-0050560116c9")])
    if(assignedBPMNModels.length != 0){
        var bpmnAssigned = "true"
    }

    //create function JSON object
    var functionJsonObject = createJsonObject(id, o_name, parent, description, hierarchyLevel, href, typeKind, typeName, bpmnAssigned, modelJsonObject)
    arrayJSON.put(functionJsonObject); //this is JSON ARRAY


    //get assigned function tree and process selection matrix

    //create additional info group VARIANT
    /*  A variant object is defined as (not the attribute value)
        If the object is connected with "is process-oriented superior" or "aggregates" or "consists of"
        and either
        the connection exist in a "Process selection matrix" model
        or the connected objects have the same 'Hierarchy Level' */
    
    //var variantObjects = getVariantObjects(objOcc, parent) //LOOP FIX IT
    
    //var variantJsonObject = createJsonObject(id, o_name, parent, description, hierarchyLevel, href, typeKind, typeName, bpmnAssigned, modelJsonObject)
    //arrayJSON.put(functionJsonObject); //this is JSON ARRAY

    //FOR THE RECURSION, NEXT LEVEL FUNC
    var assignedFuncTreeModels = objDef.AssignedModels(Constants.MT_FUNC_TREE)
    var funcObjOccs = new Array()
    for (var i = 0; i < assignedFuncTreeModels.length; i++) {
        //var tempObjDefs = assignedFuncTreeModels[i].ObjDefListFilter(Constants.OT_FUNC) //get functions for next level
        var tempObjOccs = assignedFuncTreeModels[i].ObjOccListFilter(Constants.OT_FUNC) //get functions for next level
        for (var j = 0; j < tempObjOccs.length; j++) {
            //check Hierarchy Level att
            var tempObjDef = tempObjOccs[j].ObjDef()
            if (tempObjDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("49e3ac80-b16f-11e2-424c-005056aa0ab6"), nLocale).IsMaintained()) {
                funcObjOccs.push(tempObjOccs[j])
            }
        }
    }

    //RECURSION
    for (var i = 0; i < funcObjOccs.length; i++) {
        getDataHier(funcObjOccs[i], id)
    }
}

function createJsonString(id, o_name, parent) {

    var obj = new org.json.JSONObject();
    //var obj2 = new org.json.JSONObject();

    obj.put("id", id);
    obj.put("name", o_name);
    obj.put("parent", parent);
    //obj2.put("children", children);

    //var array = new org.json.JSONArray();
    //array.put(obj2)

    //obj.put("children", array);

    return obj.toString(2);
}


function createJsonObject(id, o_name, parent, description, hierarchyLevel, href, typeKind, typeName, bpmnAssigned, objModel){

    var objMain = new org.json.JSONObject();
   
    objMain.put("id", id);
    objMain.put("name", o_name);
    objMain.put("parent", parent);
    objMain.put("description", description);
    objMain.put("hierarchyLevel", hierarchyLevel);
    objMain.put("href", href);
    objMain.put("@typeKind", typeKind);
    objMain.put("@typeName", typeName);
    objMain.put("bpmnAssigned", bpmnAssigned);
      
    objMain.put("model", objModel); //put model json in main
    
    //obj2.put("children", children);

    //var array = new org.json.JSONArray();
    //array.put(obj2)

    //obj.put("children", array);

    return objMain;
}



function createModelJsonObject(m_id, m_name, m_parent, m_level, m_status, typeName, typeKind, href, description) {

    var obj = new org.json.JSONObject();

    obj.put("id", m_id);
    obj.put("name", m_name);
    obj.put("status", m_status);
    obj.put("@typeName", typeName);
    obj.put("href", href);
    
    if(m_level != null)
    obj.put("hierarchyLevel", m_level);
    if(m_parent != null)
    obj.put("parent", m_parent);
    if(typeKind != null)
    obj.put("@typeKind", typeKind);
    if(description != null)
    obj.put("description", description);

    return obj;

}

function createVariantOfJsonObject(v_id, v_name){
    var obj = new org.json.JSONObject();
    
    obj.put("id", v_id);
    obj.put("name", v_name);
    
    return obj;
}

function createVariantJsonObject(v_id, v_name, parent, v_desc, v_href, v_att, v_typeKind, v_typeName, v_model, v_variantOf){
var objMain = new org.json.JSONObject();
   
    objMain.put("id", v_id);
    objMain.put("name", v_name);
    objMain.put("parent", parent);
    objMain.put("description", v_desc);

    objMain.put("href", v_href);
    objMain.put("variant", v_att);
    
    objMain.put("@typeKind", v_typeKind);
    objMain.put("@typeName", v_typeName);
      
    objMain.put("model", v_model); //put model json in main
    objMain.put("variantOf", v_variantOf); //put model json in main


    return objMain;

}

function existsInArray(p_obj, p_objArray) {
    var exist = false;
    for (i = 0; i < p_objArray.length; i++) {
        if (p_objArray[i].equals(p_obj)) {
            exist = true;
        }
    }
    return exist;
}

function getNestedChildren(arr, parent) {
    var out = []
    for (var i in arr) {
        if (arr[i].parent == parent) {
            var children = getNestedChildren(arr, arr[i].id)

            if (children.length) {
                arr[i].children = children
            }
            //delete(arr[i].parent)
            out.push(arr[i])
        }
    }
    return out
}

function getVariantObjects(objOcc, parent){
    var objDef = objOcc.ObjDef()
    var level = objDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("49e3ac80-b16f-11e2-424c-005056aa0ab6"), nLocale).getValue()
    
    var cxns = objDef.CxnListFilter(Constants.EDGES_INOUT, [Constants.CT_IS_PRCS_ORNT_SUPER, Constants.CT_CONS_OF_1, Constants.CT_CONS_OF_2, Constants.CT_MOD_TRANSFORM_CONSISTS_OF])
    var assignedPrcsSelMtxModels = objDef.AssignedModels(Constants.MT_PRCS_SLCT_MTX)
    
    var cxnsArray = new Array()
    for (var i in assignedPrcsSelMtxModels) {
        for (var j in cxns) {
            var cxnsOccs = cxns[j].OccList([assignedPrcsSelMtxModels[i]])
            if(cxnsOccs.length != 0){
                //cxn exist in Process selection matrix, add cxn to the array
                cxnsArray = cxnsArray.concat(cxnsOccs)
            }
        }                                      
    }
    
    //get target objects from cxnsArray
    var variantObjects = new Array()
    for (var i in cxnsArray) {
        var targetObjOcc = cxnsArray[i].getTarget()
        var sourceObjOcc = cxnsArray[i].getSource()
        if(targetObjOcc.ObjDef().GUID() != objDef.GUID()){
            variantObjects.push(targetObjOcc)
        }
        else if(sourceObjOcc.ObjDef().GUID() != objDef.GUID()){
            variantObjects.push(sourceObjOcc)
        }
    }

//*************************************************************************************************************************  
    //or the connected objects have the same 'Hierarchy Level' THIS NEEDS TO BE CHECKED -->DISABLED NEED TO FIX LOOP<--
/*     var cxnOccConected = objOcc.CxnOccList()
    for (var i in cxnOccConected) {
        var typeNum = cxnOccConected[i].getDefinition().TypeNum()
        if(typeNum == Constants.CT_IS_PRCS_ORNT_SUPER || typeNum == Constants.CT_CONS_OF_1 || typeNum == Constants.CT_CONS_OF_2 || typeNum ==Constants.CT_MOD_TRANSFORM_CONSISTS_OF){
        
        var targetObjOcc = cxnOccConected[i].getTarget()
        var sourceObjOcc = cxnOccConected[i].getSource()
        var levelTarget = targetObjOcc.ObjDef().Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("49e3ac80-b16f-11e2-424c-005056aa0ab6"), nLocale).getValue()
        var levelSource = sourceObjOcc.ObjDef().Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("49e3ac80-b16f-11e2-424c-005056aa0ab6"), nLocale).getValue()

        if(targetObjOcc.ObjDef().GUID() != objDef.GUID() && levelTarget == level){
            variantObjects.push(targetObjOcc)
        }
        else if(sourceObjOcc.ObjDef().GUID() != objDef.GUID() && levelSource == level){
            variantObjects.push(sourceObjOcc)
        }
        
        
        }
    } */
//*************************************************************************************************************************  
    variantObjects = ArisData.Unique(variantObjects)
    
    //add objects to the JSON array
    for (var i in variantObjects) {
    
    //OCCURENCE MODEL DATA
    var occModel = variantObjects[i].Model()
    var modelID = occModel.GUID()
    var modelHref = getHref(occModel, "models")
    var modelName = occModel.Name(nLocale)
    var modelStatus = occModel.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("d2ac7740-af65-11df-599b-005056aa0ab6"), nLocale).getValue()
    var modelTypeName = occModel.Type()
    //create model data JSON
    var variantModelJsonObject = createModelJsonObject(modelID, modelName, null, null, modelStatus, modelTypeName, null, modelHref, null) 
    
    var variantOfId =objDef.GUID()
    var variantOfName = objDef.Name(nLocale)
    var variantOfJsonObject = createVariantOfJsonObject(variantOfId, variantOfName)
    
    var variantName = variantObjects[i].ObjDef().Name(nLocale)
    var variantId = variantObjects[i].ObjDef().GUID()
    var variantDescription = variantObjects[i].ObjDef().Attribute(Constants.AT_DESC, nLocale).getValue()
    var variantHref = getHref(variantObjects[i].ObjDef(), "objects")
    var variantAttribute = variantObjects[i].ObjDef().Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("f1b07770-7855-11ec-7d84-00505682bf5e"), nLocale).getValue()
    var variantDepth //TODO
    var variantTypeKind = "OBJECT"
    var variantTypeName = variantObjects[i].ObjDef().Type()
    

    var variantJsonObject = createVariantJsonObject(variantId, variantName, parent, variantDescription, variantHref, variantAttribute, variantTypeKind, variantTypeName, variantModelJsonObject, variantOfJsonObject)
    arrayJSON.put(variantJsonObject); //this is JSON ARRAY
        
    }
    
    
    return variantObjects
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