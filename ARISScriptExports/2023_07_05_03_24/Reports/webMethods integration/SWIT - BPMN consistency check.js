/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

if (checkActiveFilter()) {
    var oSelModels = getSelectedBPMNModels();
    var checkMerge = new checkOpenMergeConflicts();
    if (!checkMerge.getAllMods2Open(oSelModels)){
        var mapInfoMarks = initInfoMarks(oSelModels);
        
        check_ModelForErrors( mapInfoMarks, oSelModels );
        check_Uniqueness( mapInfoMarks, oSelModels );
        
        gb_DGRMS_OPEN_SET = displayErrors();
        
        outInfoMarks( mapInfoMarks );
            
    } else {
        gb_c_ERROR_MARK_SET     = true;
        gb_DGRMS_OPEN_SET       = true;
    }       
}
writeProperties( );

/*************************************************************************************************************************************/

function checkActiveFilter() {
    var bIsFilterOk = checkFilter(ArisData.ActiveFilter());
    if (!bIsFilterOk) {
        gb_c_ERROR_MARK_SET = true;
/*    
        if (!getBoolPropertyValue("SHOW_RESULTS_WITHOUT_CALLBACK")) {
            Dialogs.MsgBox(getString("MSG_FILTER_ERROR"), Constants.MSGBOX_BTN_OK || Constants.MSGBOX_ICON_ERROR, Context.getScriptInfo(Constants.SCRIPT_NAME) );
        }
*/    
    }
    return bIsFilterOk;
}

function displayErrors() {
    if (gb_c_ERROR_MARK_SET) return true;
/*    
    if (gb_c_INFO_MARK_SET || gb_c_WARNING_MARK_SET) {
        if (getBoolPropertyValue("SHOW_RESULTS_WITHOUT_CALLBACK")) return true;
        
        dlgRes = Dialogs.MsgBox(getString("MSG_SHOW_DIAGRAMS"), Constants.MSGBOX_BTN_YESNO || Constants.MSGBOX_ICON_QUESTION, Context.getScriptInfo(Constants.SCRIPT_NAME) );
        if (dlgRes == Constants.MSGBOX_RESULT_YES) return true;
    }
*/    
    return false;
}        