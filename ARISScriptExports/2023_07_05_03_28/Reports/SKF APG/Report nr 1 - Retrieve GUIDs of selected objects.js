
var g_nLoc = 0;                                                                                             // Variable for determining the ID of the current language
var g_model_with_Selected_objects_GUID="";                                                                  // Create placeholder for the GUID of the models where the selected objects are found.
var sGUIDs="";
var g_Parent_object_to_Selected_objects_GUID= new Array();                                                                   // Create placeholders for the GUIDs of the selected objects.
var g_Parent_Object_of_Selected_objects_GUID= "";                                                           // Create placeholder for the GUID of the 'Parent' object. The Parent object is the object to which the selected objects are allocated to.
var result="OK";   
var out_GUID_model="";
var out_GUID_parent_object="";                                                                                           // The result of the report. Can be OK
function main(){
    var g_oOutFile = Context.createOutputObject(Context.getSelectedFormat(), Context.getSelectedFile());

    //g_Selected_objects_GUID.length=0;                                                                           // Reset the length of the array with selected object's GUIDs

    var i = 0;
    g_nLoc = Context.getSelectedLanguage();
    g_oOutFile.Init(g_nLoc);
    oObjOccs = ArisData.getSelectedObjOccs();

// g_Parent_Object_of_Selected_objects_GUID= ;
    g_model_with_Selected_objects_GUID=oObjOccs[0].Model().GUID();
    
    Context.setProperty(out_GUID_model, g_model_with_Selected_objects_GUID);
    

    for (i0=0;i0<oObjOccs.length;i0++){
        oCurrentObjOcc=oObjOccs[i0];
        
        
//        g_Selected_objects_GUID.push(oCurrentObjOcc.ObjDef().GUID());                                         // Push into an array the GUID of the selected object
        oParentObjects=oCurrentObjOcc.getConnectedObjOccs(1306);              // Find all objects using symbol number 1306 (= "Requirement") that are connected do the current object.
        if(oParentObjects.length>0){
            for(i1=0;i1<oParentObjects.length;i1++){
                oCurrentoParentObject=oParentObjects[i1];
                g_Parent_Object_of_Selected_objects_GUID=oCurrentoParentObject.ObjDef().GUID();                                      // Save the GUID of the parent object.
                g_Parent_object_to-Selected_objects_GUID.push(g_Parent_Object_of_Selected_objects_GUID);                        // SAbe the GUID of the parent object into an array
                
            }
        }
        else{
              result="At least one of the selected Requirements don't have a connection to a High level erquirement. Please review, correct anf try again";  
        }
    }
    
    
    
    
    sReferenceParentObjectGuid=g_Parent_object_to-Selected_objects_GUID[0];     // Pick the first of the Parent object GUIDs as reference
    for(i2=0;i2<g_Parent_object_to-Selected_objects_GUID.length;i2++){          // Compare all GUIDs of the parent GUID to the reference
        if(g_Parent_object_to-Selected_objects_GUID[i2]!=sReferenceParentObjectGuid){
            result="You have more than one Parent object in the selected content. Please review, correct, and try again.";
            
        }
        else{
        Context.setProperty(out_GUID_parent_object,sReferenceParentObjectGuid);
        }
    }
}
 
// Context.getProperty("in_GUIDs_selected_objects");    // Retrieve the list of selected object GUIDs handed over in the data flow

// HERE COMES THE MAGIC


//out_GUID_assigned_models=""
//out_GUID_obejcts_in_assigned_models=""
//out_GUID_container_model=""


//Context.setProperty(“out_GUID_assigned_models”,out_GUID_assigned_models); // GUIDS of models assigned to selected obejcts

main();