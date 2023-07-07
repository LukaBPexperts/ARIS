// Create an output report that lists all selected models with their names and names of the contained objects
// Use Name() method to ge tthe names
 
 
var oOutput = Context.createOutputObject();    // Output object
var nLocale = Context.getSelectedLanguage();   // Selected database language
 
var aModels = ArisData.getSelectedModels();    // Array of selected models
for (var i = 0; i < aModels.length; i++) {           
var oModel = aModels[i];                   // Current model
var sModelName = oModel.Name(nLocale);     // Name of current model
 
oOutput.OutputLn("Model: " + sModelName, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
   
var aObjDefs = oModel.ObjDefList();        // All object definitions that have occurrences in the model
for (var j = 0; j < aObjDefs.length; j++) {           
   var oObjDef = aObjDefs[j];             // Current object definition - increment to get every name
   var sObjName = oObjDef.Name(nLocale);  // Name of current object
       
   oOutput.OutputLn("Object: " + sObjName, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 20);
   }   
}
oOutput.WriteReport();