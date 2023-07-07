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

var g_rowData = new Array();

var g_oDatabase = ArisData.getActiveDatabase();

var g_nLoc = Context.getSelectedLanguage(); 
var g_oOutfile = Context.createOutputObject(Context.getSelectedFormat(), Context.getSelectedFile());
g_oOutfile.Init(g_nLoc);

g_oOutfile.DefineF("REPORT1", getString("TEXT21"), 24, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0, 21, 0, 0, 0, 1);
g_oOutfile.DefineF("REPORT2", getString("TEXT21"), 14, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0, 21, 0, 0, 0, 1);
g_oOutfile.DefineF("REPORT3", getString("TEXT21"), 8, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0, 21, 0, 0, 0, 1);

if ((Context.getSelectedFormat() != Constants.OutputXLS) && (Context.getSelectedFormat() != Constants.OutputXLSX)) {
    setReportHeaderFooter(g_oOutfile, g_nLoc, true, true, false);
    g_oOutfile.OutputLn("", getString("TEXT_1"), 14, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0);
    g_oOutfile.OutputLn(getString("TEXT_2"), getString("TEXT_1"), 14, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0);
}
    
g_oOutfile.BeginTable(100, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);  
g_oOutfile.TableRow();
g_oOutfile.TableCell(getString("TEXT_3"), 20, getString("TEXT_1"), 10, Constants.C_BLACK, Constants.C_LIGHT_ORANGE, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0);
g_oOutfile.TableCell(getString("TEXT_4"), 20, getString("TEXT_1"), 10, Constants.C_BLACK, Constants.C_LIGHT_ORANGE, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0);
g_oOutfile.TableCell(getString("TEXT_5"), 20, getString("TEXT_1"), 10, Constants.C_BLACK, Constants.C_LIGHT_ORANGE, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0);
g_oOutfile.TableCell(getString("TEXT_6"), 20, getString("TEXT_1"), 10, Constants.C_BLACK, Constants.C_LIGHT_ORANGE, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0);
g_oOutfile.TableCell(getString("TEXT_7"), 20, getString("TEXT_1"), 10, Constants.C_BLACK, Constants.C_LIGHT_ORANGE, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0);

var tableContent = Context.getProperty("DataFlow");
if (tableContent != null) {
    
    outDataFlow_byMacro(tableContent);
} else {

    outDataFlow_byReport();
}    

g_oOutfile.EndTable("", 100, getString("TEXT_1"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0);        
g_oOutfile.WriteReport();

function outDataFlow_byReport() {
    g_rowData = new Array();

    // Application system types selected
    var oSelApplSystemList = getSelectedApplSystems();
    
    // only Dataflow between selected application systems?
    var bOnlyBetween = false;    
    if (oSelApplSystemList.length >= 2) {
        var sQuestion = getString("TEXT_8")
        if (Dialogs.MsgBox(sQuestion, Constants.MSGBOX_ICON_QUESTION | Constants.MSGBOX_BTN_YESNO, getString("TEXT_9")) == Constants.MSGBOX_RESULT_YES) {
            bOnlyBetween = true;    
        }            
    }
    
    var aRowData = getTableContent(oSelApplSystemList, bOnlyBetween);
    var bColoredTableCell = false;        
    if (aRowData.length > 0) {
        var sSource, sTarget, sData, sProtocol, sDirection;
        
        for (var i = 0; i < aRowData.length; i++) {
            sSource = getName(aRowData[i].oSource);
            sTarget = getName(aRowData[i].oTarget);
            sData = getNames(aRowData[i].oData);
            sProtocol = getName(aRowData[i].oProtocol);
            sDirection = aRowData[i].sDirection;
            
            if ((Context.getSelectedFormat() != Constants.OutputXLS) && (Context.getSelectedFormat() != Constants.OutputXLSX)) {
                if (i > 0) {
                    if (aRowData[i].oSource.equals(aRowData[i-1].oSource)) {
                        sSource = "";
                        if (aRowData[i].oTarget.equals(aRowData[i-1].oTarget)) {  
                            sTarget = "";
                        }
                    }
                }
            }            
            g_oOutfile.TableRow();        
            g_oOutfile.TableCell(sSource, 20, getString("TEXT_1"), 10, Constants.C_BLACK, getTablecellColor(bColoredTableCell), 0, Constants.FMT_LEFT | Constants.FMT_BOLD, 0);            
            g_oOutfile.TableCell(sTarget, 20, getString("TEXT_1"), 10, Constants.C_BLACK, getTablecellColor(bColoredTableCell), 0, Constants.FMT_LEFT, 0);            
            g_oOutfile.TableCell(sData, 20, getString("TEXT_1"), 10, Constants.C_BLACK, getTablecellColor(bColoredTableCell), 0, Constants.FMT_LEFT, 0);            
            g_oOutfile.TableCell(sProtocol, 20, getString("TEXT_1"), 10, Constants.C_BLACK, getTablecellColor(bColoredTableCell), 0, Constants.FMT_LEFT, 0);            
            g_oOutfile.TableCell(sDirection, 20, getString("TEXT_1"), 10, Constants.C_BLACK, getTablecellColor(bColoredTableCell), 0, Constants.FMT_CENTER, 0);
            bColoredTableCell = !bColoredTableCell;                    
        }
    } else {
            g_oOutfile.TableRow();        
            g_oOutfile.TableCell("", 100, getString("TEXT_1"), 10, Constants.C_BLACK, getTablecellColor(bColoredTableCell), 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0);            
    }
}

function outDataFlow_byMacro(p_tableContent) {
    var bColoredTableCell = false; 

    var rowTokenizer = new java.util.StringTokenizer(p_tableContent, ";");
    while(rowTokenizer.hasMoreTokens()) {
        g_oOutfile.TableRow();
        var rowContent = rowTokenizer.nextToken();
        var nCount = 0;    
        
        columnTokenizer = new java.util.StringTokenizer(rowContent, ",");
        while(columnTokenizer.hasMoreTokens()) {
            nCount++;
            var sTableEntry = "" + columnTokenizer.nextToken();
            if (nCount <= 4) {        
                sTableEntry = getNameByGuid(sTableEntry); // nextToken = Guid of object
            }
            g_oOutfile.TableCell(sTableEntry, 20, getString("TEXT_1"), 10, Constants.C_BLACK, getTablecellColor(bColoredTableCell), 0, Constants.FMT_LEFT, 0);                        
            bColoredTableCell = !bColoredTableCell;
        }
    }
}

function getNameByGuid(p_sGuid) {
    var sName = "";
    if (p_sGuid.length > 0) {
        var oItem = g_oDatabase.FindGUID(p_sGuid);
        if (oItem.IsValid()) {
            sName = oItem.Name(g_nLoc);
        }
    }
    return sName;
}

function getSelectedApplSystems() {
    var oSelObjDefs = ArisData.getSelectedObjDefs();
    var oSelectionByType = new Array();
    for (var i = 0; i < oSelObjDefs.length; i++) {
        var oObjDef = oSelObjDefs[i];
        if (oObjDef.TypeNum() == Constants.OT_APPL_SYS_TYPE) {
            oSelectionByType.push(oObjDef);
        }
    }
    oSelectionByType.sort(sortByName);
    return oSelectionByType;
}

function getTablecellColor(p_bColored) {
    if (p_bColored)
        return Constants.C_LIGHT_YELLOW;
    
    return Constants.C_TRANSPARENT;    
}