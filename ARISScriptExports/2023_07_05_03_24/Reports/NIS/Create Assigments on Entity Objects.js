var nLocale = Context.getSelectedLanguage()
//var selectedDatabases = ArisData.getSelectedDatabases()
var selectedGroups = ArisData.getSelectedGroups()

var selectedModels = ArisData.getSelectedModels()

var selectedGroup = selectedGroups[0]


var objDefs = selectedGroup.ObjDefListFilter(Constants.OT_ENT_TYPE)

var modelList = selectedGroup.ModelList()

for (var i = 0; i < objDefs.length; i++) {

    for (var j = 0; j < modelList.length; j++) {
        if (modelList[j].Name(nLocale) == objDefs[i].Name(nLocale)) {
            //kreate assigment
            
            Context.writeStatus("Kreiranje assigmenta" + i +" od " + objDefs.length)
            
            objDefs[i].CreateAssignment( modelList[j])
        }
    }

}

