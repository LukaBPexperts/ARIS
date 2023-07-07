/* -------------
Normal variables for initating a script
------------- */
//Global variables
var g_nLoc = Context.getSelectedLanguage();
var g_exc = Context.createExcelWorkbook(Context.getSelectedFile());
var g_table = [];
//var g_entireMethodGUID = "dd838074-ac29-11d4-85b8-00005a4053ff";

/* -------------
Functions for writing to Excel
------------- */
var g_exc = Context.createExcelWorkbook(Context.getSelectedFile());
var g_headers = ["Level 0", "Level 1", "Level 2"] 
var g_columnWidths = [];
g_columnWidths.push([0, 13]);
g_columnWidths.push([1, 13]);
var g_colorGrey = getColorByRGB( 54, 53, 52);
var g_colorLighterGrey = getColorByRGB( 94, 93, 92);
var g_colorLightestGrey = getColorByRGB( 244, 243, 242);
var g_colorGreen = getColorByRGB( 143, 212, 9);


function main(){
    
    //get selection
    var oSelModels = ArisData.getSelectedModels();
    var oSelObjs = ArisData.getSelectedObjDefs();
    var oSelGroups = ArisData.getSelectedGroups();
    var oSelDatabase = ArisData.getSelectedDatabases()[0]; //Get the current database
    var oAllGroups= oSelDatabase.Find(Constants.SEARCH_GROUP) //Get all groups within the database
    
    for (var i = 0; i < oAllGroups.length; i++){
        
        //Get current group
        var oGroup = oAllGroups[i];
        var strGroupName = oGroup.Name(g_nLoc);
        
        //Get all models within the group
        var oModels = oGroup.ModelList(true, null);
        
        //Iterate over the models
        for (var j = 0; j < oModels.length; j++){
            
            //Get current model
            var oModel = oModels[j];
            var strModelName = oModel.Name(g_nLoc);
            var strModelDescription = oModel.Attribute(Constants.AT_DESC, g_nLoc).getValue()

            //Check if the current models has a value for 'Description'
            
            if (strModelDescription !== ""){

                g_table.push([strModelName, strModelDescription]);
                
                
            }
            
    
        }
        
        
    }

    
    //Writing to Excel
    printInformation(g_table, g_headers, "Model status", g_columnWidths);
    g_exc.write();
}
main();