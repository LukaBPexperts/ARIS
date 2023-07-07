var nLocale = Context.getSelectedLanguage()
//var selectedDatabases = ArisData.getSelectedDatabases()
var selectedGroups = ArisData.getSelectedGroups()

//var selectedModels = ArisData.getSelectedModels()

var selectedGroup = selectedGroups[0]
//var selectedModel = selectedModels[0]

var objDefsSlave = selectedGroup.ObjDefListFilter(Constants.OT_ERM_ATTR)

var serverName = ArisData.getActiveDatabase().ServerName()
var dbName     = ArisData.getActiveDatabase().Name( Context.getSelectedLanguage() )
var selection = Dialogs.BrowseArisItems("Select a group", "Select a group for the evaluation", serverName, dbName, Constants.CID_GROUP)


var database = ArisData.getActiveDatabase()
targetGroup = database.FindOID(selection) //pronadji tocno koja je to grupa  

var objDefsMaster = targetGroup.ObjDefListFilter(Constants.OT_ERM_ATTR)



for (var i = 0; i < objDefsMaster.length; i++) {
    
    Context.writeStatus("Konsolidacija objekta" + i +" od " + objDefsMaster.length)


    var objDefsSlave = selectedGroup.ObjDefListFilter(Constants.OT_ERM_ATTR)
    
    var objDefToCons = new Array()
    for (var j = 0; j < objDefsSlave.length; j++) {
        if (objDefsMaster[i].Name(nLocale) == objDefsSlave[j].Name(nLocale) && objDefsMaster[i].GUID() != objDefsSlave[j].GUID() ) {
            objDefToCons.push(objDefsSlave[j])
        }
    }
    

    
    
    
    
    if(objDefToCons.length != 0 && objDefToCons[0].GUID() != objDefsMaster[i].GUID() ){
    var rez = objDefsMaster[i].Consolidate(objDefToCons, true)
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