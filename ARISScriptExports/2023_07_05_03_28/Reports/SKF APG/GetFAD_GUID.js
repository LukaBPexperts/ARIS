
var Guid = Context.getProperty("ModGuid");

//var Guid = "M:02847651-1e3b-11e9-0508-00505616050b";
Guid = Guid.substring(2);

var oCurrDatabase = ArisData.getActiveDatabase();
var lng = Context.getSelectedLanguage()

var currentModel = oCurrDatabase.FindGUID(Guid)
var FAD_Guids = ";";

 var listeObjOcc = new Array()
     listeObjOcc =currentModel.ObjOccList() ;
     listeObjOcc = ArisData.sort(listeObjOcc,Constants.SORT_X,Constants.SORT_Y,lng);
     
 var cxnOccs = new Array()
     cxnOccs = currentModel.CxnOccList()
     cxnOccs = ArisData.sort(cxnOccs,Constants.SORT_X,Constants.SORT_Y,lng);
     
//    FAD_Guids = Guid;
    for(var i = 0 ; i<listeObjOcc.length; i++ ){
        var currentOcc = listeObjOcc[i] ;
        var listeModAssocie = new Array()
        listeModAssocie = listeModAssocie.concat(currentOcc.ObjDef().AssignedModels()) // All models
        for(var j = 0 ; j<listeModAssocie.length ; j++){
            var currentMod = listeModAssocie[j]    
                FAD_Guids = FAD_Guids +";M:"+ currentMod.GUID(); 
              
        } 
   
    }
    
    
    for(var i = 0 ; i<cxnOccs.length; i++ ){
        var currentCxnOcc = cxnOccs[i] ;
        var listeModAssocie = new Array()
        listeModAssocie = listeModAssocie.concat(currentCxnOcc.CxnDef().AssignedModels()) // All models
        for(var j = 0 ; j<listeModAssocie.length ; j++){
            var currentMod = listeModAssocie[j]    
                FAD_Guids = FAD_Guids +";M:"+ currentMod.GUID(); 
              
        } 
   
    }
   // remove ;;M: char
    if (FAD_Guids.indexOf(";;M:") == 0) {
        FAD_Guids = FAD_Guids.substring(2);
    }   

Context.setProperty("FADGuids",FAD_Guids);
