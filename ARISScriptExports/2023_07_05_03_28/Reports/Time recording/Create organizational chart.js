/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

Context.setProperty("excel-formulas-allowed", false); //default value is provided by server setting (tenant-specific): "abs.report.excel-formulas-allowed" 
 
//you can choose more than one epc  
var nLocale = Context.getSelectedLanguage();

var nCount_holder   = new __holder(0);
var bDelete_holder   = new __holder(false);
if (Dialog(nCount_holder, bDelete_holder)) {
    
    var oOutputFile   = Context.createOutputObject();
    var oSelFuncDefs = ArisData.getSelectedObjDefs();
    if (oSelFuncDefs.length == 0) {
        oSelFuncDefs = new Array();
        
        var oSelModels = ArisData.getSelectedModels(); //  Models
        // get relevant function objects
        for (var i = 0; i < oSelModels.length; i++) {
            var oSelFuncDefs_tmp = oSelModels[i].ObjDefListFilter(Constants.OT_FUNC);
            oSelFuncDefs = oSelFuncDefs.concat(oSelFuncDefs_tmp);
        }
    }
    
    // get relevant function objects
    var oFuncDefs_NotAssigned = new Array();
    for (var i=0; i < oSelFuncDefs.length;i++){
        var oFuncDef = oSelFuncDefs[i];
        var oOrgDefs = oFuncDef.getConnectedObjs(Constants.OT_PERS_TYPE);
        
        if (oOrgDefs.length == 0) oFuncDefs_NotAssigned.push(oFuncDef);     //list of Functions with no connection
        
        for (var j=0; j < oOrgDefs.length;j++){    
            var oOrgDef = oOrgDefs[j];
            
            if (bDelete_holder) {
                deleteModel(oOrgDef);
            }

            if (nCount_holder.value > 0) {
                //generate model
                var oAssModel = oOrgDef.Group().CreateModel(Constants.MT_ORG_CHRT, oOrgDef.Name(nLocale), nLocale);
                //...fill 
                fillModel(oAssModel, oOrgDef, nCount_holder.value);
                oOrgDef.CreateAssignment(oAssModel);
            }
        }
    }
    //Output of functions which doesn't have connections to PersTyp
    oOutputFile.BeginTable(100, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0); 
    for (var i = 0; i<oFuncDefs_NotAssigned.length; i++){
        oOutputFile.TableRow();
        oOutputFile.TableCell(oFuncDefs_NotAssigned[i].Name(nLocale), 20, "Arial", 10, Constants.C_BLACK, new java.awt.Color(0.72265625, 0.99609375, 0.72265625).getRGB(), 0, Constants.FMT_BOLD | Constants.FMT_CENTER | Constants.FMT_VTOP, 0);        
    }   
    oOutputFile.EndTable(getString("TEXT_2"), 100, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    oOutputFile.WriteReport();
}


function fillModel(oModel, oOrgDef, nCount, zahl) {
    //identify the max. sting length of entry
    var number = nCount_holder.value;
    var number2 = number.toString();
    var number3 = number2.length;
    
    var oOrgOcc = oModel.createObjOcc(oOrgDef.getDefaultSymbolNum(), oOrgDef, 1,1, true);  
    for (var i=0; i < nCount;i++){
        //identify the length of the actual number
        var aNumber = (i+1).toString();
        var aNumber2 = aNumber.length;
         
        var diff = number3 - aNumber2; //differnece between string = Number of needed " "
        var leerzeichen1 = " ";
        //var leerzeichen = diff * leerzeichen1;
        for(var e = 0; e < diff;e++){
            var leerzeichen2 = leerzeichen1+leerzeichen1;
            var leerzeichen = leerzeichen2.toString();
        }
        
        if(diff >= 1){
          var oPersTypeDef = oOrgDef.Group().CreateObjDef(Constants.OT_PERS_TYPE, oOrgDef.Name(nLocale)+" "+leerzeichen+(i+1), nLocale);
        }
        else{
          var oPersTypeDef = oOrgDef.Group().CreateObjDef(Constants.OT_PERS_TYPE, oOrgDef.Name(nLocale)+" "+(i+1), nLocale);
        }
        
        //var oPersTypeDef = oOrgDef.Group().CreateObjDef(Constants.OT_PERS_TYPE, oOrgDef.Name(nLocale)+" "+leerzeichen+(i+1), nLocale);
        var oPersTypeOcc = oModel.createObjOcc(Constants.ST_EMPL_TYPE, oPersTypeDef, 1,1, true);   

        var oCxnOcc = oModel.CreateCxnOcc(oOrgOcc, oPersTypeOcc, Constants.CT_IS_CRT_BY, getPoints(), false, false);
    }
    function getPoints() {
        var aPoints = new Array();
        aPoints.push(new java.awt.Point(0,0));
        aPoints.push(new java.awt.Point(1,1));
        return aPoints;
    }
    
    oModel.SetFlags(Constants.MODEL_LAYOUTONOPEN);
}

function deleteModel(oOrgDef) {  
    var oAssModels = oOrgDef.AssignedModels(Constants.MT_ORG_CHRT);
    for (var i = 0; i < oAssModels.length; i++) {
        var oModel = oAssModels[i];
        if (!oModel.Group().Delete(oModel)) Dialogs.MsgBox(getString("TEXT_8"))
    }
}

//Dialog
function Dialog(ndepth_holder, bDelete_holder) {
    var binput = false;   // Variable to check if input is correct
    var nuserdlg = 0;   // Variable to check if the user choose in the dialog the alternative "cancel"
    
    while (binput == false && ! (nuserdlg == 1)) {
        var userdialog = Dialogs.createNewDialogTemplate(0, 0, 300, 45, getString("TEXT_1"));
        userdialog.GroupBox(7, 0, 370, 60, getString("TEXT_3"));
        userdialog.Text(20, 15, 140, 15, getString("TEXT_4"));
        userdialog.TextBox(185, 15, 60, 21, "txtLinkLevels");
        userdialog.OKButton(60, 45, 100, 21);
        userdialog.CancelButton(180, 45, 100, 21);
        userdialog.CheckBox(40, 200, 200, 21, getString("TEXT_5"), "CHECK_DELETE");
        
        var dlg = Dialogs.createUserDialog(userdialog); 
        dlg.setDlgText("txtLinkLevels", 0);
        nuserdlg = Dialogs.show( __currentDialog = dlg);       // Show Dialog and waiting for confirmation with  OK
        
        if (isNaN(dlg.getDlgText("txtLinkLevels"))) {
            Dialogs.MsgBox(getString("TEXT_7"), Constants.MSGBOX_BTN_OK);
        } else {
            ndepth_holder.value = parseInt(dlg.getDlgText("txtLinkLevels"));
            if (ndepth_holder.value < 0) {
                Dialogs.MsgBox(getString("TEXT_6"), Constants.MSGBOX_BTN_OK);          
            } else { 
                binput = true;
            }
        }
        bDelete_holder.value = dlg.getDlgValue("CHECK_DELETE") == 1;
    }
    if (nuserdlg == 0) return false;
    return true;
}
