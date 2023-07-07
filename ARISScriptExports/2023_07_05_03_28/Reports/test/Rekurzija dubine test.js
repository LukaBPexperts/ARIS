var g_nLoc = Context.getSelectedLanguage();
var models = ArisData.getSelectedModels();
var workbook = Context.createExcelWorkbook(Context.getSelectedFile());
var dummyOutput = Context.createOutputObject();
var epcModels = new Array()  //***ANTE
//var MAXLEVEL = 2; //a maximum depth to reduce script runtime and sheet count :)
var setWrittenModels = new java.util.HashSet(); //avoid endless loops created by circular assignments
  
for(var i=0;i<models.length;i++){
    var oModel = models[i];
    createModelGraphicRecursive(oModel)
}

var a =0

function createModelGraphicRecursive(model)
{   //u globalnu varijablu epcModels spremaju se EPC-evi 
    
        if(setWrittenModels.contains(model.GUID()))
        return false; // already visited
     
    setWrittenModels.add(model.GUID())
      
    var oFuncDefs = model.ObjDefListFilter(Constants.OT_FUNC);
    
    for (var k = 0; k < oFuncDefs.length; k++) {
        var oFuncDef = oFuncDefs[k];
        var oAssignedModels = oFuncDef.AssignedModels();
         
        for (var j = 0; j < oAssignedModels.length; j++)
        {    
            var oAssModel = oAssignedModels[j];
            
            if(oAssModel.TypeNum() == Constants.MT_EEPC)
			epcModels.push(oAssignedModels[j])    //***ANTE
            
            else if (oAssModel.TypeNum() == Constants.MT_VAL_ADD_CHN_DGM){
            createModelGraphicRecursive(oAssModel)
                  
            }
        }
    }
    return true;
}
  
