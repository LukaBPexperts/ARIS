var nLocale = Context.getSelectedLanguage()
//var selectedDatabases = ArisData.getSelectedDatabases()
//var selectedGroups = ArisData.getSelectedGroups()

var selectedModels = ArisData.getSelectedModels()

//var selectedGroup = selectedGroups[0]
var selectedModel = selectedModels[0]

var occObjectModel = selectedModel.ObjOccList()

var groupModel = selectedModel.Group()

var objDefs = groupModel.ObjDefListFilter(Constants.OT_ENT_TYPE)

for (var i = 0; i < occObjectModel.length; i++) {
    
    Context.writeStatus("Konsolidacija objekta" + i +" od " + occObjectModel.length)

    
    var objDefToCons = new Array()

    for (var j = 0; j < objDefs.length; j++) {
        if (objDefs[j].Name(nLocale) == occObjectModel[i].ObjDef().Name(nLocale) && objDefs[j].GUID() != occObjectModel[i].ObjDef().GUID() ) {
            objDefToCons.push(objDefs[j])
        }
    }
    
    if(objDefToCons.length != 0 && objDefToCons[0].GUID() != occObjectModel[i].ObjDef().GUID() ){
    var rez = occObjectModel[i].ObjDef().Consolidate(objDefToCons, true)
    }

}

//konsolidacija gotova


/* for (var i = 0; i < occObjectModel.length; i++) {
    
    Context.writeStatus("Konsolidacija objekta" + i +" od " + occObjectModel.length)
    try{
    
    var objDefToCons = new Array()

    for (var j = 0; j < objDefs.length; j++) {
        if (objDefs[j].Name(nLocale) == occObjectModel[i].ObjDef().Name(nLocale) && objDefs[j].GUID() != occObjectModel[i].ObjDef().GUID() ) {
            objDefToCons.push(objDefs[j])
        }
    }
    
    if(objDefToCons.length != 0 && objDefToCons[0].GUID() != occObjectModel[i].ObjDef().GUID() ){
    var rez = occObjectModel[i].ObjDef().Consolidate(objDefToCons, true)
    }
}
catch(ex){
    var xy=0
}
} */