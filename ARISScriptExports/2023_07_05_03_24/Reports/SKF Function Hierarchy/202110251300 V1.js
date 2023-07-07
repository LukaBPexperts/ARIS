var nLocale = Context.getSelectedLanguage()
var oOutput = Context.createOutputObject()
var excelObject = Context.createExcelWorkbook(Context.getSelectedFile());


//setup excell file
//*************************************************************
var sheetName = "Sheet1"
var mainSheet = excelObject.createSheet(sheetName);
//mainSheet.setColumnWidth(0, 12000);
mainSheet.setDefaultColumnWidth(22)
var rowCounter = 0;

createHeaderRow(mainSheet)
//*************************************************************
        
       

//****************************

var g_nloc = Context.getSelectedLanguage();



function main()
{   


    var oobjdefs = ArisData.getSelectedObjDefs();
    if (oobjdefs.length > 0) {
		var ofuncs = new Array();
		
		var ncheckmsg = 0;   // Variable containing the answer to the message (2 = Abort was selected).
		var bfuncok = false;
		var binvalidfound = false; 
		
        for (var j = 0; j < oobjdefs.length; j++) {
            var ocurrentobjdef = oobjdefs[j];
            if (ocurrentobjdef.TypeNum() == Constants.OT_FUNC) {
                bfuncok = true;
                ofuncs[ofuncs.length] = ocurrentobjdef;
            } else {
                binvalidfound = true;
            }
        }
        
        if (bfuncok == false) {
            Dialogs.MsgBox(getString("TEXT8"), Constants.MSGBOX_BTN_OK, getString("TEXT4"));
            Context.setScriptError(Constants.ERR_NOFILECREATED);
        } else {
            if (binvalidfound == true) {
                ncheckmsg = Dialogs.MsgBox(getString("TEXT9"), Constants.MSGBOX_BTN_OKCANCEL, getString("TEXT4"));
            }
        }
        
        if (bfuncok == true && ! (ncheckmsg == 2)) {
            var ncount = 0; //changed from 1 to 0
            var oSuperFuncs = new Array();  
            checkhiera(ofuncs, oSuperFuncs, ncount);
            excelObject.write()
            
        } else {
            Context.setScriptError(Constants.ERR_CANCEL);
        }
    } else {
        Dialogs.MsgBox(getString("TEXT10"), Constants.MSGBOX_BTN_OKCANCEL, getString("TEXT4"));
        Context.setScriptError(Constants.ERR_NOFILECREATED);
    }
}


// ********************************************************************************************************
// *  Subprogram CheckHiera for creating the hierarchy of the functions                                *
// ********************************************************************************************************
// *  Transfer parameter                                                                                  *
// *  oObjDefs = List of functions.
// * oSuperFuncs = List of functions on higher level                                                    *
// *  nCount = counter													                            *
// ********************************************************************************************************

function checkhiera(oobjdefs, oSuperFuncs, ncount)
{
    var odoneobjdefs = new Array();	// List of objects which have already been edited on this hierarchy level.
    
    if (oobjdefs.length > 0) {
        oobjdefs = ArisData.sort(oobjdefs, Constants.AT_NAME, Constants.SORT_NONE, Constants.SORT_NONE, g_nloc);
        for (var i = 0; i < oobjdefs.length; i++) {
            
            var ocurrentobjdef = oobjdefs[i];
            
            //COMMENTED NOT USED, HEADER ALREADY DONE
            // if (ncount == 1) {
                // g_ooutfile.OutputLn("", getString("TEXT1"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
                // g_ooutfile.OutputLn((getString("TEXT11") + ocurrentobjdef.Name(g_nloc)), getString("TEXT1"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0);
            // }
            

            
            //print function in ncount column, this is level column
            var newRow = mainSheet.createRow(rowCounter);
            rowCounter++

            var newCell = newRow.createCell(ncount);
            newCell.setCellValue(ocurrentobjdef.Name(g_nloc));
            //newCell.setCellStyle(cellStyle);
            
            //g_ooutfile.OutputLn(ocurrentobjdef.Name(g_nloc), getString("TEXT1"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, nIndent);
            
            if (!isInList(ocurrentobjdef, oSuperFuncs)) {
                var oNewSuperFuncs = new Array();
                oNewSuperFuncs = oNewSuperFuncs.concat(oSuperFuncs);
                oNewSuperFuncs.push(ocurrentobjdef);
                
                var ofuncs = new Array();
                
                //COMENTED NOT USED CXN
                // First of all the direct relationships of the function.
                // var ocxns = ocurrentobjdef.CxnList();
                // if (ocxns.length > 0) {
                    // for (var h = 0; h < ocxns.length; h++) {
                        // var ocurrentcxn = ocxns[h];
                        // if ((ocurrentcxn.TypeNum() == Constants.CT_IS_PRCS_ORNT_SUPER)) {
                            // // Is process-oriented inferior.
                            // var otempobjdef = ocurrentcxn.TargetObjDef();
                            // 
                            // if (! (otempobjdef.IsEqual(ocurrentobjdef))) {
                                // var bdone = isDone(otempobjdef, odoneobjdefs);
// 
                                // if (bdone == false) {
                                    // ofuncs[ofuncs.length] = otempobjdef;
                                    // odoneobjdefs[odoneobjdefs.length] = otempobjdef;
                                // }
                            // }
                        // }
                    // }
                // }
                
                // Now the relationships that were created by the assignments.
                var omodels = ocurrentobjdef.AssignedModels();
                for (var h = 0; h < omodels.length; h++) {
                    var ocurrentmodel = omodels[h];
                    var oObjDefsInModel = ocurrentmodel.ObjDefListFilter(Constants.OT_FUNC);
                    
                    for (var k = 0; k < oObjDefsInModel.length; k++) {
                        var otempobjdef = oObjDefsInModel[k];
                        if (! (otempobjdef.IsEqual(ocurrentobjdef))) {
                            var bdone = isDone(otempobjdef, odoneobjdefs);

                            if (bdone == false) {
                                ofuncs[ofuncs.length] = otempobjdef;
                                odoneobjdefs[odoneobjdefs.length] = otempobjdef;
                            }
                        }
                    }
                }
                // Now, call recursively.
                if (ofuncs.length > 0) {
                    checkhiera(ofuncs, oNewSuperFuncs, (ncount + 1));
                }
            }
        }
    }
}

function isDone(otempobjdef, odoneobjdefs) {
    if (odoneobjdefs.length > 0) {
        for (var j = 0; j < odoneobjdefs.length; j++) {
            var ocurrentdoneobjdef = odoneobjdefs[j];
            
            if (ocurrentdoneobjdef.IsEqual(otempobjdef)) {
                return true;
            }
        }
    }
    return false;
}

function isInList(ocurrelement, oelements) {
    for (var i = 0 ; i < oelements.length; i++ ) {
        if (ocurrelement.IsEqual(oelements[i])) {
            return true;
        }
    }
    return false;
}



function setFrameStyle(outfile, iFrame) { 
    outfile.SetFrameStyle(Constants.FRAME_TOP, 0); 
    outfile.SetFrameStyle(Constants.FRAME_LEFT, 0); 
    outfile.SetFrameStyle(Constants.FRAME_RIGHT, 0); 
    outfile.SetFrameStyle(Constants.FRAME_BOTTOM, 0);
    
    outfile.SetFrameStyle(iFrame, 50, Constants.BRDR_NORMAL);
} 


main();


function createHeaderRow(headerRow) {
    var headerRow = mainSheet.createRow(rowCounter);
    rowCounter++;
    var font = excelObject.createFont();
    cellStyle1 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(169, 208, 142), Constants.SOLID_FOREGROUND);
    cellStyle2 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(155, 194, 230), Constants.SOLID_FOREGROUND);
    cellStyle2.setLocked(true)
    cellStyle3 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(255, 217, 102), Constants.SOLID_FOREGROUND);

    var newCell = headerRow.createCell(0);
    newCell.setCellValue("Model name");
    newCell.setCellStyle(cellStyle1);

    var newCell = headerRow.createCell(1);
    newCell.setCellValue("Model type");
    newCell.setCellStyle(cellStyle2);

}

function RGB(r, g, b) {
    return (new java.awt.Color(r / 255.0, g / 255.0, b / 255.0, 1)).getRGB() & 0xFFFFFF
}