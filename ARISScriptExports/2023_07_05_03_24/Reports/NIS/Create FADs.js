var nLocale = Context.getSelectedLanguage()
var selectedModels = ArisData.getSelectedModels()
var selectedOccs = ArisData.getSelectedObjOccs()

if (selectedOccs.length > 0) {
    for (var i = 0; i < selectedOccs.length; i++) {
        if (selectedOccs[i].ObjDef().TypeNum() == Constants.OT_FUNC) {
            //do creation
            var model = selectedOccs[i].Model()
            var group = model.Group()
            createFadModel(selectedOccs[i], group)
        }
    }
}
else if (selectedModels.length > 0) {
    for (var i = 0; i < selectedModels.length; i++) {
        var selectedModel = selectedModels[i]
        var selectedModelGrp = selectedModel.Group()

        //var funcObjDefs = selectedModel.ObjDefListByTypes([Constants.OT_FUNC])
        var funcObjOccs = selectedModel.ObjOccListFilter(Constants.OT_FUNC)

        for (var j = 0; j < funcObjOccs.length; j++) {
            createFadModel(funcObjOccs[j], selectedModelGrp)
        }
    }
}

function getConnectedRoleOccs(funcOcc) {
    var objOccs = funcOcc.getConnectedObjOccs(-1)

    var rolesArray = new Array()
    for (var i = 0; i < objOccs.length; i++) {
        if (objOccs[i].ObjDef().TypeNum() == Constants.OT_PERS_TYPE) {
            rolesArray.push(objOccs[i])
        }
    }
    return rolesArray
}

function getConnectedOccs(funcOcc, objType) {
    var objOccs = funcOcc.getConnectedObjOccs(-1)

    var  occsArray = new Array()
    for (var i = 0; i < objOccs.length; i++) {
        if (objOccs[i].ObjDef().TypeNum() == objType) {
            occsArray.push(objOccs[i])
        }
    }
    return occsArray
}

function createFadModel(funcObjOcc, selectedModelGrp) {
    //check if FAD already exist
    //Детаљни појединачни дијаграми 

    var fadGroup = getFadGroup(selectedModelGrp)
    
    var fadModels = fadGroup.ModelListFilter(funcObjOcc.ObjDef().Name(nLocale), nLocale, Constants.MT_FUNC_ALLOC_DGM)

    if (fadModels.length == 0) {
        //create FAD        
        var fadModel = fadGroup.CreateModel(Constants.MT_FUNC_ALLOC_DGM, funcObjOcc.ObjDef().Name(nLocale), nLocale)
        var assignmet = funcObjOcc.ObjDef().CreateAssignment(fadModel, false)
        var funcFadOcc = fadModel.createObjOcc(funcObjOcc.getSymbol(), funcObjOcc.ObjDef(), 200, 200, true)

        //create role occs
        //var roleObjOccs = getConnectedRoleOccs(funcObjOcc)
        var roleObjOccs = getConnectedOccs(funcObjOcc, Constants.OT_PERS_TYPE)
        createObjOccsOnModel(funcFadOcc, roleObjOccs, fadModel, Constants.CT_EXEC_2)
        
        /* for (var k = 0; k < roleObjOccs.length; k++) {
            try{
            var roleX = 400
            var roleY = 400
            var roleFadOcc = fadModel.createObjOcc(roleObjOccs[k].getSymbol(), roleObjOccs[k].ObjDef(), roleX, roleY, true)
            var pointList = new Array()
            var oPointSource = new java.awt.Point(funcFadOcc.X(), funcFadOcc.Y());
            var oPointTarget = new java.awt.Point(roleFadOcc.X(), roleFadOcc.Y());
            pointList.push(oPointSource)
            pointList.push(oPointTarget)
            var tempCxnOcc = fadModel.CreateCxnOcc(roleFadOcc, funcFadOcc, Constants.CT_EXEC_2, pointList)
            }catch(e){}
        } */
        
        var orgUnitObjOccs = getConnectedOccs(funcObjOcc, Constants.OT_ORG_UNIT_TYPE)
        createObjOccsOnModel(funcFadOcc, orgUnitObjOccs, fadModel, Constants.CT_EXEC_2)
        
        var layoutParameters = ArisData.getModelOptions().getLayoutParameters(fadModel.TypeNum())
        layoutParameters.setArrangeSatellites(1) //square
        layoutParameters.setInsertSpace(true)
        layoutParameters.setSpacingX(100)
        layoutParameters.setSpacingY(100)
        layoutParameters.setMinimizeCxnAnchors(true)

        var layoutRes = fadModel.doLayout()
    }
    else if (fadModels.length == 1) {
        //fad already exist
        var fadModel = fadModels[0]
    }
    else {
        //error, couple of models with the same name already exist
    }
}

function getFadGroup(selectedModelGrp){
    var childGroups = selectedModelGrp.Childs()
 
    for(var i=0; i<childGroups.length; i++){
        if(childGroups[i].Name(nLocale) == "Детаљни појединачни дијаграми"){
            return childGroups[i]
        }       
    }  	
var newGroup = selectedModelGrp.CreateChildGroup("Детаљни појединачни дијаграми", nLocale)
return newGroup
}

function createObjOccsOnModel(funcFadOcc, objOccs, model, cxnType){
    for (var k = 0; k < objOccs.length; k++) {
        try{
        var objX = 400
        var objY = 400
        var objFadOcc = model.createObjOcc(objOccs[k].getSymbol(), objOccs[k].ObjDef(), objX, objY, true)
        var pointList = new Array()
        var oPointSource = new java.awt.Point(funcFadOcc.X(), funcFadOcc.Y());
        var oPointTarget = new java.awt.Point(objFadOcc.X(), objFadOcc.Y());
        pointList.push(oPointSource)
        pointList.push(oPointTarget)
        var tempCxnOcc = model.CreateCxnOcc(objFadOcc, funcFadOcc, cxnType, pointList)
        }catch(e){}
    }      
}