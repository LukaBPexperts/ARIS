
var Guid = Context.getProperty("ModGuid");

//var Guid = "M:cb382861-a65e-11e5-38c7-080027650386";
Guid = Guid.substring(2);

var oCurrDatabase = ArisData.getActiveDatabase();
var lng = Context.getSelectedLanguage()

var currentModel = oCurrDatabase.FindGUID(Guid)
var FAD_Guids = ";";

 var listeObjOcc = new Array()
     listeObjOcc =currentModel.ObjOccList() ;
     listeObjOcc = ArisData.sort(listeObjOcc,Constants.SORT_X,Constants.SORT_Y,lng);
     
//    FAD_Guids = Guid;
    for(var i = 0 ; i<listeObjOcc.length; i++ ){
        var currentOcc = listeObjOcc[i] ;
        var listeModAssocie = new Array()
        listeModAssocie = listeModAssocie.concat(currentOcc.ObjDef().AssignedModels(14)) // MDL_FAD
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
