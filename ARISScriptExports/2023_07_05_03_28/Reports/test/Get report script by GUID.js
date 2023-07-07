var nLoc = Context.getSelectedLanguage();
const componentID   = 1; //COMP_REPORT

var scriptAdmin = Context.getComponent("ScriptAdmin");

var aCategories = scriptAdmin.getCategories(componentID, nLoc)
    
for (var i in aCategories) {
    var category = aCategories[i];
    var aScriptInfos = scriptAdmin.getScriptInfos(componentID, category.getCategoryID(), nLoc);
    
    if (aScriptInfos.length > 0) {
        for (var j in aScriptInfos) {
            var scriptInfo = aScriptInfos[j];

            var scriptName =scriptInfo.getName()
            var scriptGUID = scriptInfo.getID()
            
            var a = 2
        }
    }
}