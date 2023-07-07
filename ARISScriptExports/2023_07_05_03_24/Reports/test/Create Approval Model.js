//---------------------------------------------------------------------------------------------------------------
// ARIS Script - Create Approval Model
//
// (c) Ingo Opetnik
// Software GmbH Austria
//
// Version 1.0.0 (01.12.2015|Ingo Opetnik)
//	- Basic functionality
//
//----------------------------------------------------------------------------------------------------------------  
    
    var g_nLoc = Context.getSelectedLanguage();
    var g_hmTypeMapping = {};
    var g_sGroupPath = Context.getProperty("in_folder");
    //var g_sGroupPath = ArisData.getActiveDatabase().RootGroup().Path(-1);
    
    var g_sInActualCR = Context.getProperty("in_actual_cr");
    //var g_sInActualCR = "10";

    var g_sModelGUID = "";
    var g_sStatus = "OK";
   
    main();
    
// ------------------------------------------------------------------------------------------------------------------------
//* Routine to configure object type to model type
// ------------------------------------------------------------------------------------------------------------------------
function INIT_Configuration() {
    g_hmTypeMapping[Constants.OT_APPL_SYS_CLS] = Constants.MT_ACS_DGM; // /Main group/B Content Library/Applications
    g_hmTypeMapping[Constants.OT_ORG_UNIT] = Constants.MT_ACS_DGM; // /Main group/B Content Library/Applications
    g_hmTypeMapping[Constants.OT_STRCT_ELMT] = Constants.MT_ACS_DGM; // /Main group/B Content Library/Applications
    g_hmTypeMapping[Constants.OT_APPL_SYS_TYPE] = Constants.MT_ACS_DGM; // /Main group/B Content Library/Applications    
    // g_hmTypeMapping[Constants.OT_SYS_ORG_UNIT] = Constants.MT_FUNC_ALLOC_DGM;
    
    
    
//    g_hmTypeMapping.put(Constants.OT_FUNC, '0bfec820-b400-11de-209d-005056a70192');             // /Main group/B Content Library/Interfaces (VAC L1)

//    g_hmTypeMapping.put(Constants.OT_FUNC, Constants.MT_VAL_ADD_CHN_DGM);             // /Main group/B Content Library/Interfaces
    
    //g_hmTypeMapping.put(Constants.OT_FUNC, Constants.MT_FUNC_ALLOC_DGM);             // /Main group/B Content Library/Interfaces    
    //g_hmTypeMapping.put(Constants.OT_INFO_CARR, Constants.MT_FUNC_ALLOC_DGM);             // /Main group/B Content Library/Interfaces        
    
//    g_hmTypeMapping.put(Constants.OT_FUNC, Constants.MT_EEPC);             // /Main group/B Content Library/Interfaces
    
//    g_hmTypeMapping.put(Constants.OT_FUNC, Constants.MT_BPMN_COLLABORATION_DIAGRAM);              // /Main group/B Content Library/Organizational Units/Legal Units AND /Main group/B Content Library/Organizational Units/Operating Units
    
    //g_hmTypeMapping.put(Constants.OT_LOC, Constants.MT_ORG_CHRT);                   // /Main group/B Content Library/Organizational Units/Location    
}

//------------------------------------------------------------------------------------------------------------------------
//------------------------------------------------------------------------------------------------------------------------
//------------------------------------------------------------------------------------------------------------------------
//------------------------------------------------------------------------------------------------------------------------
//------------------------------------------------------------------------------------------------------------------------

//*
//* 'Main' routine of the script
//*
function main() {
    var aoSelObjDefs = new Array();   
    
    aoSelObjDefs = ArisData.getSelectedObjDefs();
    
    createApprovalModel(aoSelObjDefs);
    
    Context.setProperty("out_model_guid", g_sModelGUID);
    Context.setProperty("out_status", g_sStatus);
}

function createApprovalModel(aoObjDefs) {
    
    INIT_Configuration();
       
    var nModelTypeNum = g_hmTypeMapping[aoObjDefs[0].TypeNum()];   
    if (nModelTypeNum == null) {
        nModelTypeNum = Constants.MT_FUNC_ALLOC_DGM;
    }
    
    // remove first / char
    if (g_sGroupPath.indexOf("/") == 0) {
        g_sGroupPath = g_sGroupPath.substring(1);
    }
    
    var oGroup = ArisData.getActiveDatabase().Group(g_sGroupPath, g_nLoc);  
    if (oGroup == null || !oGroup.IsValid()) {
        g_sStatus = "Could not find group '" + g_sGroupPath + "'!";
        return;    
    }
    
    var oModel = createModel(g_sInActualCR + " - " + oGroup.Name(g_nLoc), nModelTypeNum, oGroup); 
    if (oModel == null || !oModel.IsValid()) {
        g_sStatus = "Could not create model!";
        return;    
    }
    
    aoObjDefs = ArisData.sort(aoObjDefs, Constants.AT_NAME, g_nLoc);
    for (var i = 0; i < aoObjDefs.length; i++) {  
        var objOcc = null;
        try {
            objOcc = oModel.createObjOcc(aoObjDefs[i].getDefaultSymbolNum(), aoObjDefs[i], 10, 10 + 10*i, true, true);
        } catch (e) {}
        if (objOcc==null) {
            g_sStatus = "Could not create symbol with number " + aoObjDefs[i].getDefaultSymbolNum() + " on a model type " + nModelTypeNum + "!";
        }
    }
    oModel.ApplyTemplate();
    oModel.doLayout();   

    g_sModelGUID = "M:" + oModel.GUID();
}                
    
    
function createModel(sName, nTypeNum, oGroup) {    
    return oGroup.CreateModel(nTypeNum, sName, g_nLoc);                
}    

// FOR TESTING
//     Context.setProperty("out_model_guid", "M:f41ac371-877b-11e5-6b2c-08002766f7ea");