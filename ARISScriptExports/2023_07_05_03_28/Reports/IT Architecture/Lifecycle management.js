/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

 // Report Configuration
const ALLOW_COLOR_OBJECTS_IN_CONNECT = false;   // Allow dialog options 'Color objects', "Reset colored objects" - Currently not supported in ARIS Connect (BLUE-12275)

/****************************************************/ 

var cFONT_SIZE = 8

// ----------------------------------------------------------
//
// evaluate getString("TEXT_15"):  cEVAL_ID_LC = "LC"      
// evaluate getString("TEXT_16"):  cEVAL_ID_SC = "SC";      
//
// ----------------------------------------------------------

Context.setProperty("excel-formulas-allowed", false); //default value is provided by server setting (tenant-specific): "abs.report.excel-formulas-allowed"

var aOT_LifeCycle = new Array();
aOT_LifeCycle.push(Constants.OT_APPL_SYS);
aOT_LifeCycle.push(Constants.OT_HW_CMP);

var aOT_StdCycle = new Array();
aOT_StdCycle.push(Constants.OT_APPL_SYS_TYPE);
aOT_StdCycle.push(Constants.OT_DBMS_TYPE);
aOT_StdCycle.push(Constants.OT_PRG_LNG);
aOT_StdCycle.push(Constants.OT_HW_CMP_TYPE);
aOT_StdCycle.push(Constants.OT_OS_TYPE);
aOT_StdCycle.push(Constants.OT_FUNC);
aOT_StdCycle.push(Constants.OT_TECH_TRM);
aOT_StdCycle.push(Constants.OT_PERF);
aOT_StdCycle.push(Constants.OT_NW_PROT);

var g_nLoc = Context.getSelectedLanguage(); 
var g_oOutfile = null;

function MACRO_OBJOCC_COLOR(p_oObjOcc, p_nColor) {
    this.oObjOcc    = p_oObjOcc;
    this.nColor     = p_nColor;
}

function MACRO_LEGEND_COLOR(p_sState, p_nColor) {
    this.sState     = p_sState;
    this.nColor     = p_nColor;
}


// needed, if report was started by MACRO
var g_aObjColors_Macro = new Array();
var g_aLegend_LC = new Array();
var g_aLegend_SC = new Array();

var g_bIsStartedByMacro = getBoolPropertyValue("isStartedByMacro");  // Anubis 541446
var g_bLegend_LC = false;
var g_bLegend_SC = false;

main();

// needed, if report was started by MACRO
if (g_bIsStartedByMacro) {
    macro_writeProperties();
    
    if (g_bLegend_LC) macro_writeLegend_LC();
    if (g_bLegend_SC) macro_writeLegend_SC();
}

/*----------------------------------------------------------------------------*/

function main() {
/*
    // Initialized in 'itArchitecture.js'
    var g_sDate_Ref = getStringDate(new Date());    
    var g_sDate_Start = "01.01.2018";
    var g_sDate_End   = "31.12.2022";
*/    
    var bOutTable_holder = new __holder(false);
    var nSetObjColor_holder = new __holder(0);
    var bAddLegend_holder = new __holder(false);    
    var nState_holder = new __holder(0);        
    var dDate_Ref = new Date();

    if (dlgSelectOptions(bOutTable_holder, nSetObjColor_holder, bAddLegend_holder, nState_holder)) {

        g_nDate_Ref = getLongDate(g_sDate_Ref);
        g_nDate_Start = getLongDate(g_sDate_Start);
        g_nDate_End   = getLongDate(g_sDate_End);
        
        // --- Out Table ---
        if (bOutTable_holder.value) {
            g_oOutfile = Context.createOutputObject(Context.getSelectedFormat(), Context.getSelectedFile());
            g_oOutfile.Init(g_nLoc);
            
            // Set Landscape
            var pageHeight = g_oOutfile.GetPageHeight();
            var pageWidth = g_oOutfile.GetPageWidth();
            if (pageHeight > pageWidth) {
                g_oOutfile.SetPageHeight(pageWidth);
                g_oOutfile.SetPageWidth(pageHeight);
            }            
            
        if ( (Context.getSelectedFormat() != Constants.OutputXLS) && (Context.getSelectedFormat() != Constants.OutputXLSX) )
            setReportHeaderFooter(g_oOutfile, g_nLoc, false, false, false);

        
// --- Lebenzyklus ---
            var oObjDefList_LC = getSelectedObjDefList(cEVAL_ID_LC);        // get 'selected' object definitions
            if (oObjDefList_LC.length > 0) {
                var sTitle = getString("TEXT_15")
                outHead(sTitle);      
                g_oOutfile.BeginTable(100, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_LEFT, 0);
                g_oOutfile.TableRow();
                g_oOutfile.TableCell(getString("TEXT_1"), 13, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_BOLD | Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
                g_oOutfile.TableCell(getString("TEXT_3"), 7, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_BOLD | Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
                g_oOutfile.TableCell(getString("TEXT_4"), 10, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, cCOLOR_LC_Plan, 0, Constants.FMT_BOLD | Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
                g_oOutfile.TableCell(getString("TEXT_5"), 8, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, cCOLOR_LC_Procure, 0, Constants.FMT_BOLD | Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
                g_oOutfile.TableCell(getString("TEXT_6"), 8, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, cCOLOR_LC_Develop, 0, Constants.FMT_BOLD | Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
                g_oOutfile.TableCell(getString("TEXT_7"), 8, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, cCOLOR_LC_Test, 0, Constants.FMT_BOLD | Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
                g_oOutfile.TableCell(getString("TEXT_8"), 8, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, cCOLOR_LC_Operation, 0, Constants.FMT_BOLD | Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
                g_oOutfile.TableCell(getString("TEXT_9"), 8, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, cCOLOR_LC_Deact, 0, Constants.FMT_BOLD | Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
                
                g_oOutfile.TableCell("", 30, getString("TEXT_2"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0);
                var oArrow = getArrowPicture(getString("TEXT_2"));
                g_oOutfile.OutGraphic(oArrow, -1, getCellWidth(g_oOutfile, 30)-2, 10);

                var bColoredTableCell = false;
                for ( i = 0 ; i < oObjDefList_LC.length ; i++ ) {
                    outObject(oObjDefList_LC[i], (nState_holder.value == 0), cEVAL_ID_LC, bColoredTableCell);
                    bColoredTableCell = !bColoredTableCell;                    
                }
                if ((Context.getSelectedFormat() != Constants.OutputXLS) && (Context.getSelectedFormat() != Constants.OutputXLSX))
                    sTitle = "";
                g_oOutfile.EndTable(sTitle, 100, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT, 0);
            }

// --- Standardisierungszyklus ---                        
            var oObjDefList_SC = getSelectedObjDefList(cEVAL_ID_SC);          // get 'selected' object definitions
            if (oObjDefList_SC.length > 0) {
                var sTitle = getString("TEXT_16")
                outHead(sTitle);      
                
                g_oOutfile.BeginTable(100, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_LEFT, 0);
                g_oOutfile.TableRow();
                g_oOutfile.TableCell(getString("TEXT_1"), 16, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_BOLD | Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
                g_oOutfile.TableCell(getString("TEXT_3"), 7, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_BOLD | Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
                g_oOutfile.TableCell(getString("TEXT_10"), 11, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, cCOLOR_SC_Eval, 0, Constants.FMT_BOLD | Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
                g_oOutfile.TableCell(getString("TEXT_11"), 9, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, cCOLOR_SC_Request, 0, Constants.FMT_BOLD | Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
                g_oOutfile.TableCell(getString("TEXT_12"), 9, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, cCOLOR_SC_PhasedIn, 0, Constants.FMT_BOLD | Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
                g_oOutfile.TableCell(getString("TEXT_13"), 9, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, cCOLOR_SC_Standard, 0, Constants.FMT_BOLD | Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
                g_oOutfile.TableCell(getString("TEXT_14"), 9, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, cCOLOR_SC_PhasedOut, 0, Constants.FMT_BOLD | Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);

                g_oOutfile.TableCell("", 30, getString("TEXT_2"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0);
                var oArrow = getArrowPicture(getString("TEXT_2"));
                g_oOutfile.OutGraphic(oArrow, -1, getCellWidth(g_oOutfile, 30)-2, 10);
                
                var bColoredTableCell = false;
                for ( i = 0 ; i < oObjDefList_SC.length ; i++ ) {
                    outObject(oObjDefList_SC[i], (nState_holder.value == 0), cEVAL_ID_SC, bColoredTableCell);
                    bColoredTableCell = !bColoredTableCell;                    
                }
                if ((Context.getSelectedFormat() != Constants.OutputXLS) && (Context.getSelectedFormat() != Constants.OutputXLSX))
                    sTitle = "";
                g_oOutfile.EndTable(sTitle, 100, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT, 0);
            }
            
            if (bAddLegend_holder.value) {
                var LCtext;
                if (oObjDefList_LC.length > 0) LCtext = getString("TEXT_15");

                var SCtext;
                if (oObjDefList_SC.length > 0) SCtext = getString("TEXT_16");
                
                var oLegendPicture = getLegendPicture(LCtext, SCtext, getString("TEXT_2"));
                
                if ((Context.getSelectedFormat() != Constants.OutputXLS) && (Context.getSelectedFormat() != Constants.OutputXLSX)) {
                    g_oOutfile.OutputLn("", getString("TEXT_2"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
                    g_oOutfile.OutGraphic(oLegendPicture, -1, 500, 500);
                } else {
                    g_oOutfile.BeginTable(100, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_LEFT, 0);
                    g_oOutfile.TableRow();
                    g_oOutfile.TableCell("", 100, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_BOLD | Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
                    g_oOutfile.OutGraphic(oLegendPicture, -1, 100, 100);
                    g_oOutfile.EndTable(getString("TEXT_31"), 100, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT, 0);
                }
            }
            
            g_oOutfile.WriteReport(Context.getSelectedPath(), Context.getSelectedFile());            
        } else {
            Context.setScriptError(Constants.ERR_NOFILECREATED);
        }
        // --- Set Colors ---
        if (nSetObjColor_holder.value != 0) {
// --- Lebenzyklus ---            
            var oObjOccList_LC = getSelectedObjOccList(cEVAL_ID_LC);          // get 'selected' object occurences
            for (var i = 0 ; i < oObjOccList_LC.length ; i++ ) {  
                setObjectColor(oObjOccList_LC[i], (nSetObjColor_holder.value > 0), (nState_holder.value == 0), cEVAL_ID_LC);
            }
// --- Standardisierungszyklus ---            
            var oObjOccList_SC = getSelectedObjOccList(cEVAL_ID_SC);          // get 'selected' object occurences
            for (var i = 0 ; i < oObjOccList_SC.length ; i++ ) {  
                setObjectColor(oObjOccList_SC[i], (nSetObjColor_holder.value > 0), (nState_holder.value == 0), cEVAL_ID_SC);
            }            
            if (bAddLegend_holder.value) {
                if (oObjOccList_LC.length > 0) g_bLegend_LC = true;
                if (oObjOccList_SC.length > 0) g_bLegend_SC = true;
                
                if (!g_bIsStartedByMacro) {
                    addLegend(oObjOccList_LC, oObjOccList_SC);
                }
            }
        }
    } else { Context.setScriptError(Constants.ERR_CANCEL) }
}

function getSelectedObjDefList(p_EvalID) {
    // set array of object types
    var aObjTypeNums = new Array();
    if (p_EvalID == cEVAL_ID_LC) {
        aObjTypeNums = aOT_LifeCycle;
    } else {
        aObjTypeNums = aOT_StdCycle;
    }
    
    var oSelObjDefs = new Array();

    var mList = ArisData.getSelectedModels();
    for (var i = 0; i < mList.length; i++) {
        var oModel = mList[i];
        for (var j = 0; j < aObjTypeNums.length; j++) {
            oSelObjDefs = oSelObjDefs.concat(oModel.ObjDefListFilter(aObjTypeNums[j]));
        }
    }
    var oList = ArisData.getSelectedObjDefs();
    for (var i = 0; i < oList.length; i++) {
        var oObjDef = oList[i];
        for (var j = 0; j < aObjTypeNums.length; j++) {
            if (oObjDef.TypeNum() == aObjTypeNums[j]) {
                oSelObjDefs.push(oObjDef);
                break;
            }
        }
    }
    oSelObjDefs = ArisData.Unique(oSelObjDefs);
    oSelObjDefs.sort(new ArraySortComparator(Constants.AT_NAME, Constants.SORT_TYPE, Constants.SORT_NONE, g_nLoc).compare);
    return oSelObjDefs;
}

function getSelectedObjOccList(p_EvalID) {
    // set array of object types
    var aObjTypeNums = new Array();
    if (p_EvalID == cEVAL_ID_LC) {
        aObjTypeNums = aOT_LifeCycle;
    } else {
        aObjTypeNums = aOT_StdCycle;
    }

    var oSelObjOccs = new Array();

    var mList = ArisData.getSelectedModels();
    for (var i = 0; i < mList.length; i++) {
        var oModel = mList[i];
        for (var j = 0; j < aObjTypeNums.length; j++) {
            oSelObjOccs = oSelObjOccs.concat(oModel.ObjOccListFilter(aObjTypeNums[j]));
        }
    }
    var oList = ArisData.getSelectedObjOccs();
    for (var i = 0; i < oList.length; i++) {
        var oObjOcc = oList[i];
        for (var j = 0; j < aObjTypeNums.length; j++) {
            if (oObjOcc.ObjDef().TypeNum() == aObjTypeNums[j]) {
                oSelObjOccs.push(oObjOcc);
                break;
            }
        }
    }
    return filterOutOccsInPSM(oSelObjOccs);
    
    function filterOutOccsInPSM(oObjOccs) {        
        var oFilteredObjOccs = new Array();
        for (var i = 0; i < oObjOccs.length; i++) {
            var oObjOcc = oObjOccs[i];
            if (oObjOcc.Model().TypeNum() != Constants.MT_SYS_LAY_OUT_PLAN) {
                oFilteredObjOccs.push(oObjOcc);
            }
        }
        return oFilteredObjOccs;
    }
}

function setObjectColor(p_oObjOcc, p_bSetColor, p_bColorOfState, p_EvalID) {
    var nColor = cCOLOR_DEFAULT;    /* = -1 */
    // bSetColor == true -> set color, else -> reset color
    if (p_bSetColor) {
        if (p_bColorOfState) {
            nColor = getColorOfState(p_oObjOcc.ObjDef(), p_EvalID);        
        } else {
            nColor = getColorOfPhase(p_oObjOcc.ObjDef(), p_EvalID);
        }
    }
    if (!g_bIsStartedByMacro) {
        try {
            if (nColor == cCOLOR_DEFAULT) {
                p_oObjOcc.setDefaultProperties(true, true, false);      // Set default properties
            } else {
                p_oObjOcc.setColor(nColor);                             // Set color of occs
            }
        } catch (e) {}
    } else {
        // needed, if report was started by MACRO
        g_aObjColors_Macro.push(new MACRO_OBJOCC_COLOR(p_oObjOcc, new java.lang.Long(nColor)));
    }
}
    
function outObject(p_oObjDef, p_bColorOfState, p_EvalID, p_bColoredTableCell) {

    g_oOutfile.TableRow();
    if (p_EvalID == cEVAL_ID_LC) {
// --- Lebenzyklus ---    
        g_oOutfile.TableCell(p_oObjDef.Name(g_nLoc), 13, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, getTablecellColor(p_bColoredTableCell), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VCENTER, 0);
    } else {     
        g_oOutfile.TableCell(p_oObjDef.Name(g_nLoc), 16, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, getTablecellColor(p_bColoredTableCell), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VCENTER, 0);
    }           
    var oStatePicture = getStatePicture(p_oObjDef, p_bColorOfState, p_EvalID);    
    g_oOutfile.TableCell("", 7, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, getTablecellColor(p_bColoredTableCell), 0, Constants.FMT_CENTER, 0);
    g_oOutfile.OutGraphic(oStatePicture, -1, getCellWidth(g_oOutfile, 7), 10);
//    g_oOutfile.OutGraphic(oStatePicture, -1, 23, 10);    

    if (p_EvalID == cEVAL_ID_LC) {
// --- Lebenzyklus ---    
        var dDate_LC_Plan_Start     = getValueOfDateAttr(p_oObjDef, cAT_LC_Plan_Start);
        var dDate_LC_Plan_End       = getValueOfDateAttr(p_oObjDef, cAT_LC_Plan_End);
        var dDate_LC_Procure_End    = getValueOfDateAttr(p_oObjDef, cAT_LC_Procure_End);
        var dDate_LC_Develop_End    = getValueOfDateAttr(p_oObjDef, cAT_LC_Develop_End);
        var dDate_LC_Test_End       = getValueOfDateAttr(p_oObjDef, cAT_LC_Test_End);
        var dDate_LC_Operation_End  = getValueOfDateAttr(p_oObjDef, cAT_LC_Operation_End);
        var dDate_LC_Deact_End      = getValueOfDateAttr(p_oObjDef, cAT_LC_Deact_End);

        g_oOutfile.TableCell(dDate_LC_Plan_Start + " -\n" + dDate_LC_Plan_End, 10, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, getTablecellColor(p_bColoredTableCell), 0, Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
        g_oOutfile.TableCell(dDate_LC_Procure_End, 8, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, getTablecellColor(p_bColoredTableCell), 0, Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
        g_oOutfile.TableCell(dDate_LC_Develop_End, 8, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, getTablecellColor(p_bColoredTableCell), 0, Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
        g_oOutfile.TableCell(dDate_LC_Test_End, 8, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, getTablecellColor(p_bColoredTableCell), 0, Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
        g_oOutfile.TableCell(dDate_LC_Operation_End, 8, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, getTablecellColor(p_bColoredTableCell), 0, Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
        g_oOutfile.TableCell(dDate_LC_Deact_End, 8, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, getTablecellColor(p_bColoredTableCell), 0, Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
    } else {     
        var dDate_SC_Eval_Start     = getValueOfDateAttr(p_oObjDef, cAT_SC_Eval_Start);        
        var dDate_SC_Eval_End       = getValueOfDateAttr(p_oObjDef, cAT_SC_Eval_End);
        var dDate_SC_Request_End    = getValueOfDateAttr(p_oObjDef, cAT_SC_Request_End);
        var dDate_SC_PhasedIn_End   = getValueOfDateAttr(p_oObjDef, cAT_SC_PhasedIn_End);
        var dDate_SC_Standard_End   = getValueOfDateAttr(p_oObjDef, cAT_SC_Standard_End);
        var dDate_SC_PhasedOut_End  = getValueOfDateAttr(p_oObjDef, cAT_SC_PhasedOut_End);

        g_oOutfile.TableCell(dDate_SC_Eval_Start + " -\n" + dDate_SC_Eval_End, 11, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, getTablecellColor(p_bColoredTableCell), 0, Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
        g_oOutfile.TableCell(dDate_SC_Request_End, 9, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, getTablecellColor(p_bColoredTableCell), 0, Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
        g_oOutfile.TableCell(dDate_SC_PhasedIn_End, 9, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, getTablecellColor(p_bColoredTableCell), 0, Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
        g_oOutfile.TableCell(dDate_SC_Standard_End, 9, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, getTablecellColor(p_bColoredTableCell), 0, Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
        g_oOutfile.TableCell(dDate_SC_PhasedOut_End, 9, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, getTablecellColor(p_bColoredTableCell), 0, Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
    }
    var oGraphicPicture = getGraphicPicture(p_oObjDef, p_EvalID, getString("TEXT_17"), getString("TEXT_2"));
    g_oOutfile.TableCell("", 30, getString("TEXT_2"), cFONT_SIZE, Constants.C_BLACK, getTablecellColor(p_bColoredTableCell), 0, Constants.FMT_LEFT, 0);
    g_oOutfile.OutGraphic(oGraphicPicture, -1, getCellWidth(g_oOutfile, 30)-2, 10);
//    g_oOutfile.OutGraphic(oGraphicPicture, -1, 70, 10);    
}

function dlgSelectOptions(p_bOutTable_holder, p_nSetObjColor_holder, p_bAddLegend_holder, p_nState_holder) {
    var nuserdlg = 0;   // Variable for the user dialog box
    
    var userdialog = Dialogs.createNewDialogTemplate(500, 280, getString("TEXT_18"), "dlgFuncSelectOptions");
    userdialog.GroupBox(10, 10, 450, 85, getString("TEXT_19"));
    userdialog.CheckBox(30, 25, 400, 15, getString("TEXT_20"), "var_outTable");
    userdialog.CheckBox(30, 40, 400, 15, getString("TEXT_21"), "var_setColor");
    userdialog.CheckBox(30, 55, 400, 15, getString("TEXT_22"), "var_resetColor");
    userdialog.CheckBox(30, 70, 400, 15, getString("TEXT_23"), "var_addLegend");
    
    userdialog.GroupBox(10, 100, 450, 55, getString("TEXT_24"));
    userdialog.OptionGroup("var_states");
    userdialog.OptionButton(30, 115, 350, 15, getString("TEXT_25"));
    userdialog.OptionButton(30, 130, 350, 15, getString("TEXT_26"));
       
    userdialog.GroupBox(10, 160, 450, 140, getString("TEXT_27"));
    userdialog.Text(30, 177, 250, 14, getString("TEXT_26") + ":");
    userdialog.DateChooser(300, 175, 120, 21, "var_dateRef");
    userdialog.GroupBox(20, 205, 430, 70, getString("TEXT_28"));
    userdialog.Text(30, 222, 250, 14, getString("TEXT_29") + ":");
    userdialog.DateChooser(300, 220, 120, 21, "var_dateStart");
    userdialog.Text(30, 247, 250, 14, getString("TEXT_30") + ":");
    userdialog.DateChooser(300, 245, 120, 21, "var_dateEnd");

    userdialog.OKButton();
    userdialog.CancelButton();
    userdialog.HelpButton("HID_2551a890_1a2d_11da_5bb8_000802c68187_dlg_01.hlp");
    
    var dlg = Dialogs.createUserDialog(userdialog); 
    var sSection = "SCRIPT_2551a890_1a2d_11da_5bb8_000802c68187";
    
    var bExitDialog = false;
    while (!bExitDialog) {
        // Read dialog settings from config    
        ReadSettingsDlgValue(dlg, sSection, "var_outTable", 1);
        ReadSettingsDlgValue(dlg, sSection, "var_setColor", 0);
        ReadSettingsDlgValue(dlg, sSection, "var_resetColor", 0);
        ReadSettingsDlgValue(dlg, sSection, "var_addLegend", 0);
        ReadSettingsDlgValue(dlg, sSection, "var_states", 0);
        ReadSettingsDlgText(dlg, sSection, "var_dateRef", convertToInternal(g_sDate_Ref));
        ReadSettingsDlgText(dlg, sSection, "var_dateStart", convertToInternal(g_sDate_Start));
        ReadSettingsDlgText(dlg, sSection, "var_dateEnd", convertToInternal(g_sDate_End));  
        
        nuserdlg = Dialogs.show( __currentDialog = dlg);    // Showing dialog and waiting for confirmation with OK
        
        if (nuserdlg == 0) {
            bExitDialog = true;
        } else {
            if ((getDate("" + convertFromInternal(dlg.getDlgText("var_dateRef"))) != null) &&
                (getDate("" + convertFromInternal(dlg.getDlgText("var_dateStart"))) != null) &&
                (getDate("" + convertFromInternal(dlg.getDlgText("var_dateEnd"))) != null)) {
                    
                bExitDialog = true;
                
                p_bOutTable_holder.value = (dlg.getDlgValue("var_outTable") == 1);
                
                if (dlg.getDlgValue("var_setColor") == 1) {
                    p_nSetObjColor_holder.value = 1;
                }
                if (dlg.getDlgValue("var_resetColor") == 1) {
                    p_nSetObjColor_holder.value = -1;
                }
                
                p_nState_holder.value = dlg.getDlgValue("var_states");                
                p_bAddLegend_holder.value = (dlg.getDlgValue("var_addLegend") == 1);    // Add legend
                    
                g_sDate_Ref   = "" + convertFromInternal(dlg.getDlgText("var_dateRef"));
                g_sDate_Start = "" + convertFromInternal(dlg.getDlgText("var_dateStart"));
                g_sDate_End   = "" + convertFromInternal(dlg.getDlgText("var_dateEnd"));
            }
        }
    }
    // Write dialog settings to config    
    if (nuserdlg != 0) {
        WriteSettingsDlgValue(dlg, sSection, "var_outTable");
        WriteSettingsDlgValue(dlg, sSection, "var_setColor");
        WriteSettingsDlgValue(dlg, sSection, "var_resetColor");
        WriteSettingsDlgValue(dlg, sSection, "var_addLegend");
        WriteSettingsDlgValue(dlg, sSection, "var_states");
        WriteSettingsDlgText(dlg, sSection, "var_dateRef");
        WriteSettingsDlgText(dlg, sSection, "var_dateStart");
        WriteSettingsDlgText(dlg, sSection, "var_dateEnd");  
    }
    return (nuserdlg != 0);
}

function dlgFuncSelectOptions(dlgitem, action, suppvalue) {
    var result = false;
    
    switch(action) {
        case 1:
            if (__currentDialog.getDlgValue("var_states") != 1) {
                __currentDialog.setDlgEnable("var_dateRef", false);
            }          
            
            if (!Context.getEnvironment().equals(Constants.ENVIRONMENT_STD) && !ALLOW_COLOR_OBJECTS_IN_CONNECT) {
                // BLUE-12275 Disable dialog options 'Color objects', "Reset colored objects" , because currently not supported in ARIS Connect
                __currentDialog.setDlgValue("var_setColor", 0);
                __currentDialog.setDlgValue("var_resetColor", 0);
                __currentDialog.setDlgEnable("var_setColor", false);
                __currentDialog.setDlgEnable("var_resetColor", false);

            } else {
                if (ArisData.getSelectedModels().length > 0 || ArisData.getSelectedObjOccs().length > 0) {
                    __currentDialog.setDlgEnable("var_setColor", true);
                    __currentDialog.setDlgEnable("var_resetColor", true);
                    
                    if (__currentDialog.getDlgValue("var_setColor") != 0) {
                        __currentDialog.setDlgValue("var_resetColor", 0);
                        __currentDialog.setDlgEnable("var_resetColor", false);
                    } else {
                        if (__currentDialog.getDlgValue("var_resetColor") != 0) {
                            __currentDialog.setDlgEnable("var_setColor", false);
                        }
                    }
                } else {
                    __currentDialog.setDlgEnable("var_setColor", false);
                    __currentDialog.setDlgEnable("var_resetColor", false);
                }
            }
            break;
        case 2:
            switch(dlgitem) {
                case "var_outTable":                
                    if (__currentDialog.getDlgValue("var_outTable") == 0) {
                        __currentDialog.setDlgEnable("var_dateStart", false);
                        __currentDialog.setDlgEnable("var_dateEnd", false);
                    } else {
                        __currentDialog.setDlgEnable("var_dateStart", true);
                        __currentDialog.setDlgEnable("var_dateEnd", true);
                    }
                    result = true;
                    break;
                case "var_setColor":
                    if (__currentDialog.getDlgValue("var_setColor") == 0) {
                        __currentDialog.setDlgEnable("var_resetColor", true);
                    } else {
                        __currentDialog.setDlgEnable("var_resetColor", false);
                    }
                    result = true;
                    break;
                case "var_states":
                    if (__currentDialog.getDlgValue("var_states") == 0) {
                        __currentDialog.setDlgEnable("var_dateRef", false);
                    } else {
                        __currentDialog.setDlgEnable("var_dateRef", true);
                    }
                    result = true;
                    break;
                case "var_resetColor":
                    if (__currentDialog.getDlgValue("var_resetColor") == 0) {
                        __currentDialog.setDlgEnable("var_setColor", true);
                    } else {
                        __currentDialog.setDlgEnable("var_setColor", false);
                    }
                    result = true;
                    break;
            }
    }
    return result;
}

function addLegend(p_oObjOccList_LC, p_oObjOccList_SC) {
    var oObjOccList = new Array();
    oObjOccList = oObjOccList.concat(p_oObjOccList_LC);
    oObjOccList = oObjOccList.concat(p_oObjOccList_SC);
    
    var oModelList = new Array();
    for (var i = 0 ; i < oObjOccList.length ; i++ ) {  
        oModelList.push(oObjOccList[i].Model());
    }            
    oModelList = ArisData.Unique(oModelList);

    for (var i = 0 ; i < oModelList.length ; i++ ) {  
        var oModel = oModelList[i];
        try {
            oModel.doLayout();
        } catch (e) {}
        var yStart = getMaxYPos(oModel) + 100;
        var xStart = 100;

        if (g_bLegend_LC) {
            // Lebenzyklus
            oModel.CreateTextOcc (xStart, yStart + 50, getString("TEXT_15"));
            createColoredRect(oModel, xStart, yStart + 100, cCOLOR_LC_Plan,      getNameOfAttrValue(cATV_LC_State_Plan));
            createColoredRect(oModel, xStart, yStart + 200, cCOLOR_LC_Procure,   getNameOfAttrValue(cATV_LC_State_Procure));
            createColoredRect(oModel, xStart, yStart + 300, cCOLOR_LC_Develop,   getNameOfAttrValue(cATV_LC_State_Develop));
            createColoredRect(oModel, xStart, yStart + 400, cCOLOR_LC_Test,      getNameOfAttrValue(cATV_LC_State_Test));
            createColoredRect(oModel, xStart, yStart + 500, cCOLOR_LC_Operation, getNameOfAttrValue(cATV_LC_State_Operation));
            createColoredRect(oModel, xStart, yStart + 600, cCOLOR_LC_Deact,     getNameOfAttrValue(cATV_LC_State_Deact));
            createColoredRect(oModel, xStart, yStart + 700, cCOLOR_LC_Offline,   getNameOfAttrValue(cATV_LC_State_IsDeact));
            
            xStart = 600;
        }
        
        if (g_bLegend_SC) {
            // Standardisierungszyklus
            oModel.CreateTextOcc (xStart, yStart + 50, getString("TEXT_16"));        
            createColoredRect(oModel, xStart, yStart + 100, cCOLOR_SC_NonStandard, getNameOfAttrValue(cATV_SC_State_NonStandard));        
            createColoredRect(oModel, xStart, yStart + 200, cCOLOR_SC_Eval,        getNameOfAttrValue(cATV_SC_State_Eval));
            createColoredRect(oModel, xStart, yStart + 300, cCOLOR_SC_Request,     getNameOfAttrValue(cATV_SC_State_Request));
            createColoredRect(oModel, xStart, yStart + 400, cCOLOR_SC_PhasedIn,    getNameOfAttrValue(cATV_SC_State_PhasedIn));
            createColoredRect(oModel, xStart, yStart + 500, cCOLOR_SC_Standard,    getNameOfAttrValue(cATV_SC_State_Standard));
            createColoredRect(oModel, xStart, yStart + 600, cCOLOR_SC_Std_Ltd,     getNameOfAttrValue(cATV_SC_State_Std_Ltd));
            createColoredRect(oModel, xStart, yStart + 700, cCOLOR_SC_PhasedOut,   getNameOfAttrValue(cATV_SC_State_PhasedOut));
            createColoredRect(oModel, xStart, yStart + 800, cCOLOR_SC_NonStandard, getNameOfAttrValue(cATV_SC_State_IsPhasedOut));                
            createColoredRect(oModel, xStart, yStart + 900, cCOLOR_SC_Refused,     getNameOfAttrValue(cATV_SC_State_Refused));
        }
    }    
}

function getMaxYPos(p_oModel) {
    var yMax = 0;
    var oObjOccList = p_oModel.ObjOccList();
    for (var i = 0 ; i < oObjOccList.length ; i++ ) {
        var oObjOcc = oObjOccList[i];
        var yTmp = oObjOcc.Y() + oObjOcc.Height();
        
        if (yTmp > yMax) 
            yMax = yTmp;
    }  
    return yMax; 
} 
 
function createColoredRect(p_oModel, p_xPos, p_yPos, p_nColor, p_sState) {
    var oRect = p_oModel.createRoundedRectangle(p_xPos, p_yPos, 100, 50);
    if (oRect != null) {
        oRect.setRoundness(0,0);
        oRect.setBrushColor(p_nColor);
        
        p_oModel.CreateTextOcc (p_xPos + 150, p_yPos + 25, p_sState);
    }
} 

function macro_writeProperties() {
    if (g_aObjColors_Macro.length > 0) {
        var sb = new java.lang.StringBuffer();
        
        for ( i = 0 ; i < g_aObjColors_Macro.length ; i++ ) {
            sb.append(g_aObjColors_Macro[i].oObjOcc.ObjectID());
            sb.append(",");
            sb.append(g_aObjColors_Macro[i].nColor);
            sb.append(";");
        }
        Context.setProperty("LifeCycle", sb.toString());   
    }
}

function macro_writeLegend_LC() {
//  Lebenzyklus    
    var sb = new java.lang.StringBuffer();
    sb.append("" +  cCOLOR_LC_Plan          + "," + getNameOfAttrValue(cATV_LC_State_Plan) + ";");
    sb.append("" +  cCOLOR_LC_Procure       + "," + getNameOfAttrValue(cATV_LC_State_Procure) + ";");
    sb.append("" + cCOLOR_LC_Develop        + "," + getNameOfAttrValue(cATV_LC_State_Develop) + ";");
    sb.append("" + cCOLOR_LC_Test           + "," + getNameOfAttrValue(cATV_LC_State_Test) + ";");
    sb.append("" + cCOLOR_LC_Operation      + "," + getNameOfAttrValue(cATV_LC_State_Operation) + ";");
    sb.append("" + cCOLOR_LC_Deact          + "," + getNameOfAttrValue(cATV_LC_State_Deact) + ";");
    sb.append("" + cCOLOR_LC_Offline        + "," + getNameOfAttrValue(cATV_LC_State_IsDeact) + ";");
    Context.setProperty("Legend_LC", sb.toString());       
}

function macro_writeLegend_SC() {
//  Standardisierungszyklus    
    var sb = new java.lang.StringBuffer();
    sb.append("" +  cCOLOR_SC_NonStandard   + "," + getNameOfAttrValue(cATV_SC_State_NonStandard) + ";");
    sb.append("" +  cCOLOR_SC_Eval          + "," + getNameOfAttrValue(cATV_SC_State_Eval) + ";");
    sb.append("" +  cCOLOR_SC_Request       + "," + getNameOfAttrValue(cATV_SC_State_Request) + ";");
    sb.append("" +  cCOLOR_SC_PhasedIn      + "," + getNameOfAttrValue(cATV_SC_State_PhasedIn) + ";");
    sb.append("" +  cCOLOR_SC_Standard      + "," + getNameOfAttrValue(cATV_SC_State_Standard) + ";");
    sb.append("" +  cCOLOR_SC_Std_Ltd       + "," + getNameOfAttrValue(cATV_SC_State_Std_Ltd) + ";");
    sb.append("" +  cCOLOR_SC_PhasedOut     + "," + getNameOfAttrValue(cATV_SC_State_PhasedOut) + ";");
    sb.append("" +  cCOLOR_SC_NonStandard   + "," + getNameOfAttrValue(cATV_SC_State_IsPhasedOut) + ";");
    sb.append("" +  cCOLOR_SC_Refused       + "," + getNameOfAttrValue(cATV_SC_State_Refused) + ";");
    Context.setProperty("Legend_SC", sb.toString());       
}

function getTablecellColor(p_bColored) {
    if (p_bColored)
        return Constants.C_LIGHT_YELLOW;
    
    return Constants.C_TRANSPARENT;    
}

function outHead(p_sText) {
    if ((Context.getSelectedFormat() != Constants.OutputXLS) && (Context.getSelectedFormat() != Constants.OutputXLSX)) {
        g_oOutfile.OutputLn("", getString("TEXT_2"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);    
        g_oOutfile.Output(p_sText + ": ", getString("TEXT_2"), 14, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_BOLD, 0);
        if (ArisData.getSelectedModels().length == 1) {
            // Anubis 436425
            g_oOutfile.Output(ArisData.getSelectedModels()[0].Name(g_nLoc), getString("TEXT_2"), 14, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_BOLD, 0);
        }        
        g_oOutfile.OutputLn("", getString("TEXT_2"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
    }
}

function getBoolPropertyValue(p_sPropKey) {
    var property = Context.getProperty(p_sPropKey);
    if (property != null) {
        return (StrComp(property, "true") == 0);
    }
    return false;
}


