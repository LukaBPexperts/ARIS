/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

// BLUE-16474 Additional output of attributes of UML elements

Context.setProperty("excel-formulas-allowed", false); //default value is provided by server setting (tenant-specific): "abs.report.excel-formulas-allowed"

function CHECK_TYPE() {
    this.ndatabase = 0;
    this.ngroup = 0;
    this.nmodel = 0;
    this.nobject = 0;
    this.ncxn = 0;
    this.nshortcut = 0;
    this.ntable = 0;
    this.ntext = 0;
    this.nfontstyle = 0;
    this.nlanguage = 0;
    this.nsourcelang = 0; 
    this.ntargetlang = 0; 
    this.brecursive = false; 
    this.bwithtarget = false;   // only if source language <> target language then output of target language
    this.bonlynotmaint = false;
    this.bexpand = false;       // BLUE-11810 Additional output of attributes which are maintained only in the target language
}

var g_tcheck = new CHECK_TYPE();


var g_omethodfilter = ArisData.ActiveFilter();
var g_nloc = Context.getSelectedLanguage();
var g_ooutfile = Context.createOutputObject();
g_ooutfile.Init(g_nloc);

var g_bColored = false; // global variable to change background color of table rows

var g_nwidth = 0; 
var g_nrowmax = 0; 
var g_nwidth = 30;
var g_nrowmax = 60000;


main();


/* ------------------------------------------------------------------------------------ */

function main() {

    var bformat = false; 
    if (Context.getSelectedFormat() == Constants.OutputXLS || Context.getSelectedFormat() == Constants.OutputXLSX) {     // BLUE-10385 Support XLSX format
        bformat = true;
    } else {
        if ( Dialogs.MsgBox(getString("TEXT1"), Constants.MSGBOX_BTN_YESNOCANCEL | Constants.MSGBOX_ICON_QUESTION, getString("TEXT2")) == Constants.MSGBOX_RESULT_YES) {    // BLUE-10385
            Context.setSelectedFormat(Constants.OutputXLS);
            Context.setSelectedFile(changeextension(Context.getSelectedFile(), "xls"));
            g_ooutfile = Context.createOutputObject();            
            bformat = true;
        }
    }
    
    if (bformat == true) {
        
        if (userdlg(g_tcheck)) {
            // iterate through oSelectedGroups
            var oselectedgroups = ArisData.getSelectedGroups();
            if (oselectedgroups.length > 0) {
                
                var bRootGroupRecursive = isSelectedRootGroupAndRecursive(oselectedgroups);
                
                // out database attributes
                if (g_tcheck.ndatabase == 1) {
                    Context.writeOutput(getString("TEXT3"));
                    
                    var odatabaselist = new Array();
                    odatabaselist.push(ArisData.getActiveDatabase());
                    var text = getString("TEXT12");
                    outitemlist(odatabaselist, false, getString("TEXT12"), 1);
                }
                // out group attributes
                if (g_tcheck.ngroup == 1) {
                    Context.writeOutput(getString("TEXT4"));

                    var ogrouplist = new Array();
                    if (bRootGroupRecursive) {
                        ogrouplist = ArisData.getActiveDatabase().Find(Constants.SEARCH_GROUP);                        
                    } else {
                        for (var i = 0 ; i < oselectedgroups.length ; i++ ) {
                            var ocurrgroup = oselectedgroups[i];
                            ogrouplist.push(ocurrgroup);
                            ogrouplist = ogrouplist.concat(filterArisGroups( ocurrgroup.Childs(g_tcheck.brecursive) ));     // BLUE-16474
                        }
                    }
                    outitemlist(ogrouplist, false, getString("TEXT14"), 0);
                }
                // out model attributes
                if (g_tcheck.nmodel == 1) {
                    Context.writeOutput(getString("TEXT5"));
                    
                    var omodellist = new Array();
                    if (bRootGroupRecursive) {
                        omodellist = ArisData.getActiveDatabase().Find(Constants.SEARCH_MODEL);
                    } else {
                        for (var i = 0 ; i < oselectedgroups.length ; i++ ) {
                            var ocurrgroup = oselectedgroups[i];
                            omodellist = omodellist.concat(ocurrgroup.ModelList(g_tcheck.brecursive));
                        }
                    }
                    outitemlist(omodellist, true, getString("TEXT15"), 0);
                }
                // out object attributes
                if (g_tcheck.nobject == 1) {
                    Context.writeOutput(getString("TEXT6"));

                    var oobjectlist = new Array();
                    if (bRootGroupRecursive) {
                        oobjectlist = ArisData.getActiveDatabase().Find(Constants.SEARCH_OBJDEF);
                    } else {
                        for (var i = 0 ; i < oselectedgroups.length ; i++ ) {
                            var ocurrgroup = oselectedgroups[i];
                            oobjectlist = oobjectlist.concat(ocurrgroup.ObjDefList(g_tcheck.brecursive));
                        }
                    }
                    outitemlist(oobjectlist, true, getString("TEXT16"), 0);
                }
                
                // out UML element attributes (BLUE-16474)
                if (g_tcheck.nobject == 1) {
                    
                    var oUmlElements = getUmlElements(bRootGroupRecursive, oselectedgroups);
                    outitemlist(oUmlElements, true, getString("TEXT43"), 0);
                }
                
                // out cxn attributes
                if (g_tcheck.ncxn == 1) {
                    Context.writeOutput(getString("TEXT7"));
                    
                    var ocxnlist = new Array();
                    if (bRootGroupRecursive) {
                        ocxnlist = ArisData.getActiveDatabase().Find(Constants.SEARCH_CXNDEF);
                    } else {
                        for (var i = 0 ; i < oselectedgroups.length ; i++ ) {
                            var ocurrgroup = oselectedgroups[i];
                            var oobjectlist = ocurrgroup.ObjDefList(g_tcheck.brecursive);
                            
                            for (var j = 0 ; j < oobjectlist.length; j++ ) {
                                ocxnlist = ocxnlist.concat(oobjectlist[j].CxnList(Constants.EDGES_INOUT));
                            }
                        }
                        ocxnlist = ArisData.Unique(ocxnlist);       // UNIQUE !
                    }
                    outcxnlist(ocxnlist, getString("TEXT17"));
                }
                // out shortcut attributes
                if (g_tcheck.nshortcut == 1) {
                    Context.writeOutput(getString("TEXT39"));
                    
                    var oshortcutlist = new Array();
                    if (bRootGroupRecursive) {
                        oshortcutlist = ArisData.getActiveDatabase().Find(Constants.SEARCH_SHORTCUT);
                    } else {
                        for (var i = 0 ; i < oselectedgroups.length ; i++ ) {
                            var ocurrgroup = oselectedgroups[i];
                            var ogrouplist = new Array();
                            ogrouplist.push(ocurrgroup);
                            ogrouplist = ogrouplist.concat(ocurrgroup.Childs(g_tcheck.brecursive));
        
                            for (var j = 0 ; j < ogrouplist.length ; j++ ) {
                                oshortcutlist = oshortcutlist.concat(ogrouplist[j].Shortcuts(0, false));
                            }                        
                        }
                    }
                    outitemlist(oshortcutlist, false, getString("TEXT40"), 0);
                }
                // out text definition attributes
                if (g_tcheck.ntext == 1) {
                    Context.writeOutput(getString("TEXT8"));
                    
                    var otextdeflist = getTextDefs();                                                   // BLUE-12155
                    outitemlist(otextdeflist, false, getString("TEXT18"), 0);
                }
                // out fontstyle attributes
                if (g_tcheck.nfontstyle == 1) {
                    Context.writeOutput(getString("TEXT10"));

                    var ofontstylelist = ArisData.getActiveDatabase().FontStyleList();
                    outitemlist(ofontstylelist, false, getString("TEXT20"), 0);
                }
                g_ooutfile.WriteReport(Context.getSelectedPath(), Context.getSelectedFile());
            } else {
                Dialogs.MsgBox(getString("TEXT11"), Constants.MSGBOX_BTN_OK, getString("TEXT2"));       // BLUE-10385
                Context.setScriptError(Constants.ERR_CANCEL);
            }
        } else {
            Context.setScriptError(Constants.ERR_CANCEL);
        }
    }  else {
        Context.setScriptError(Constants.ERR_CANCEL);
    }
}

function getUmlElements(bRootGroupRecursive, oselectedgroups) {
    var aUml2ObjTypes = ArisData.ActiveFilter().getMetamodelItems(2/*UML2*/, Constants.CID_OBJDEF);
    if (aUml2ObjTypes.length == 0) return new Array();

    var oUmlElements = filterUmlElements( ArisData.getActiveDatabase().Find(Constants.SEARCH_GROUP, aUml2ObjTypes) );
    if (bRootGroupRecursive) {                        
        return oUmlElements;
    } else {
        return filterSelected(oUmlElements);
    }
       
    function filterSelected(oUmlElements) {
        var selectedUmlElements = new Array()
        for (var i in oUmlElements) {
            var oUmlElement = oUmlElements[i];
            
            if (isSelectedElement(oUmlElement)) {
                selectedUmlElements.push(oUmlElement);
            }
        }
        return selectedUmlElements;
        
        function isSelectedElement(oUmlElement) {
            if (g_tcheck.brecursive) {
                // UML elements in selected groups and their child groups
                for (var i in oselectedgroups) {
                    if (oUmlElement.IsChildGroupOf(oselectedgroups[i])) {
                        return true;
                    } 
                }
            } else {
                // Only UML elements in selected groups (without child groups)
                var oParent = getParentGroup(oUmlElement);
                for (var i in oselectedgroups) {
                    if (oParent.IsEqual(oselectedgroups[i])) {
                        return true;
                    } 
                }
            }
            return false;
        }
        
        function getParentGroup(oUmlElement) {
            var oParent = oUmlElement.Parent();
            if (oParent.IsValid()) {
                if (isArisGroup(oParent)) return oParent;
                
                return getParentGroup(oParent);
            }
            return null;
        }
    }
}

function isArisGroup(oGroup) { return oGroup.TypeNum() == 9999 }

function filterArisGroups(oGroups)  { return filterArisGroupsInt(oGroups, true) }
function filterUmlElements(oGroups) { return filterArisGroupsInt(oGroups, false) }

function filterArisGroupsInt(oGroups, bFilterThis) {
    var oFilteredGroups = new Array();
    for (var i in oGroups) {
        var oGroup = oGroups[i]
        if (isArisGroup(oGroup) == bFilterThis) {
            oFilteredGroups.push(oGroup);
        }
    }
    return oFilteredGroups;
}

function getTextDefs() {
    // BLUE-12155 Remove all entries in this list where the attribute AT_MODEL_AT is maintained
    var oTextDefs_filtered = new Array();
    var oTextDefs = ArisData.getActiveDatabase().TextDefList();
    for (var i=0; i<oTextDefs.length; i++) {
        var oTextDef = oTextDefs[i];
        if (!oTextDef.Attribute(Constants.AT_MODEL_AT, g_nloc).IsMaintained()) oTextDefs_filtered.push(oTextDef);
    }
    return oTextDefs_filtered;
}

function isSelectedRootGroupAndRecursive(p_oselectedgroups) {
    if (p_oselectedgroups.length > 0 && g_tcheck.brecursive) {
        if (p_oselectedgroups[0].IsEqual(ArisData.getActiveDatabase().RootGroup())) {
            return true;
        }
    }
    return false;
}

function outitemlist(oitemlist, bwithtype, ssheetname, nguidentry) {    
    var ncount_holder = new __holder(1);
    var nrow_holder = new __holder(2);
    
    if (oitemlist.length > 0) {
        g_ooutfile.BeginTable(100, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
        outitemheadlines(getItemKind(oitemlist[0]), bwithtype);
        
        g_bColored = false;        
        for (var i = 0 ; i < oitemlist.length ; i++ ) {
            var oitem = oitemlist[i];
            // Output of item attributes
            outitemattributes(oitem, oitem.Name(g_nloc), bwithtype, ssheetname, ncount_holder, nrow_holder, nguidentry);
        }
        
        if (ncount_holder.value > 1) {
            ssheetname = ssheetname + "_" + ncount_holder.value;
        }
        g_ooutfile.EndTable(ssheetname, 100, getString("TEXT13"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    }
}

function outcxnlist(ocxnlist, ssheetname) {
    var ncount_holder = new __holder(1);
    var nrow_holder = new __holder(2);
    
    if (ocxnlist.length > 0) {
        g_ooutfile.BeginTable(100, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
        outcxnheadlines();
        
        g_bColored = false;
        for (var i = 0 ; i < ocxnlist.length ; i++ ) {
            var ocurrcxn = ocxnlist[i];
            // Output of cxn attributes
            outcxnattributes(ocurrcxn, ssheetname, ncount_holder, nrow_holder);
        }
        
        if (ncount_holder.value > 1) {
            ssheetname = ssheetname + "_" + ncount_holder.value;
        }
        g_ooutfile.EndTable(ssheetname, 100, getString("TEXT13"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    }
}

function getItemKind(item) {
    // AKC-7973 Add item kind information to output -> More performant search in import-report (f86fc130-eaea-11d8-12e0-9d2843560f51)
    var nKindNum = item.KindNum();
    switch (nKindNum) {
        case Constants.CID_GROUP:
        case Constants.CID_MODEL:
        case Constants.CID_OBJDEF:
        case Constants.CID_CXNDEF:
            return nKindNum;
    }
    return null;
}


function outitemheadlines(nKindNum, bwithtype) {
    var sKindNum = (nKindNum != null) ? nKindNum : "";  // AKC-7973
    
    g_ooutfile.TableRow();
    g_ooutfile.TableCell(getString("TEXT21"), g_nwidth, getString("TEXT13"), 10, Constants.C_WHITE, getTableCellColor_Bk(true), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    
    if (bwithtype) {
        g_ooutfile.TableCell(getString("TEXT22"), g_nwidth, getString("TEXT13"), 10, Constants.C_WHITE, getTableCellColor_Bk(true), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    }
    g_ooutfile.TableCell(getString("TEXT23"), g_nwidth, getString("TEXT13"), 10, Constants.C_WHITE, getTableCellColor_Bk(true), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.TableCell(getString("TEXT24"), g_nwidth, getString("TEXT13"), 10, Constants.C_WHITE, getTableCellColor_Bk(true), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.TableCell(getString("TEXT25"), g_nwidth, getString("TEXT13"), 10, Constants.C_WHITE, getTableCellColor_Bk(true), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.TableCell(getString("TEXT26"), g_nwidth, getString("TEXT13"), 10, Constants.C_WHITE, getTableCellColor_Bk(true), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    
    if (g_tcheck.bwithtarget) {
        g_ooutfile.TableCell(getString("TEXT27"), g_nwidth, getString("TEXT13"), 10, Constants.C_WHITE, getTableCellColor_Bk(true), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    }
    g_ooutfile.TableRow();
    g_ooutfile.TableCell("GUID", g_nwidth, getString("TEXT13"), 10, Constants.C_GREY_50_PERCENT, getTableCellColor_Bk(true), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    
    if (bwithtype) {
        g_ooutfile.TableCell("", g_nwidth, getString("TEXT13"), 10, Constants.C_GREY_50_PERCENT, getTableCellColor_Bk(true), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    }
    g_ooutfile.TableCell(sKindNum, g_nwidth, getString("TEXT13"), 10, Constants.C_GREY_50_PERCENT, getTableCellColor_Bk(true), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.TableCell("AttrType", g_nwidth, getString("TEXT13"), 10, Constants.C_GREY_50_PERCENT, getTableCellColor_Bk(true), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.TableCell("", g_nwidth, getString("TEXT13"), 10, Constants.C_GREY_50_PERCENT, getTableCellColor_Bk(true), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.TableCell(g_tcheck.nsourcelang, g_nwidth, getString("TEXT13"), 10, Constants.C_GREY_50_PERCENT, getTableCellColor_Bk(true), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    
    if (g_tcheck.bwithtarget) {
        g_ooutfile.TableCell(g_tcheck.ntargetlang, g_nwidth, getString("TEXT13"), 10, Constants.C_GREY_50_PERCENT, getTableCellColor_Bk(true), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    }
}

function outcxnheadlines() {
    var sKindNum = Constants.CID_CXNDEF;  // AKC-7973
    
    g_ooutfile.TableRow();
    g_ooutfile.TableCell(getString("TEXT21"), g_nwidth, getString("TEXT13"), 10, Constants.C_WHITE, getTableCellColor_Bk(true), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.TableCell(getString("TEXT28"), g_nwidth, getString("TEXT13"), 10, Constants.C_WHITE, getTableCellColor_Bk(true), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.TableCell(getString("TEXT29"), g_nwidth, getString("TEXT13"), 10, Constants.C_WHITE, getTableCellColor_Bk(true), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.TableCell(getString("TEXT30"), g_nwidth, getString("TEXT13"), 10, Constants.C_WHITE, getTableCellColor_Bk(true), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.TableCell(getString("TEXT31"), g_nwidth, getString("TEXT13"), 10, Constants.C_WHITE, getTableCellColor_Bk(true), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.TableCell(getString("TEXT25"), g_nwidth, getString("TEXT13"), 10, Constants.C_WHITE, getTableCellColor_Bk(true), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.TableCell(getString("TEXT26"), g_nwidth, getString("TEXT13"), 10, Constants.C_WHITE, getTableCellColor_Bk(true), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    
    if (g_tcheck.bwithtarget) {
        g_ooutfile.TableCell(getString("TEXT27"), g_nwidth, getString("TEXT13"), 10, Constants.C_WHITE, getTableCellColor_Bk(true), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    }
    g_ooutfile.TableRow();
    g_ooutfile.TableCell("GUID", g_nwidth, getString("TEXT13"), 10, Constants.C_GREY_50_PERCENT, getTableCellColor_Bk(true), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.TableCell(sKindNum, g_nwidth, getString("TEXT13"), 10, Constants.C_GREY_50_PERCENT, getTableCellColor_Bk(true), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.TableCell("", g_nwidth, getString("TEXT13"), 10, Constants.C_GREY_50_PERCENT, getTableCellColor_Bk(true), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.TableCell("", g_nwidth, getString("TEXT13"), 10, Constants.C_GREY_50_PERCENT, getTableCellColor_Bk(true), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.TableCell("AttrType", g_nwidth, getString("TEXT13"), 10, Constants.C_GREY_50_PERCENT, getTableCellColor_Bk(true), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.TableCell("", g_nwidth, getString("TEXT13"), 10, Constants.C_GREY_50_PERCENT, getTableCellColor_Bk(true), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.TableCell(g_tcheck.nsourcelang, g_nwidth, getString("TEXT13"), 10, Constants.C_GREY_50_PERCENT, getTableCellColor_Bk(true), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    
    if (g_tcheck.bwithtarget) {
        g_ooutfile.TableCell(g_tcheck.ntargetlang, g_nwidth, getString("TEXT13"), 10, Constants.C_GREY_50_PERCENT, getTableCellColor_Bk(true), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    }
}

function outitemattributes(ocurritem, sname, bwithtype, ssheetname, ncount_holder, nrow_holder, nguidentry) {
    var ssheetnametmp = ssheetname;

    var sguid = "";    
    switch(nguidentry) {
        case 0:
            sguid = ocurritem.GUID();
            break;
        case 1:
            sguid = "Database_" + ocurritem.Name(g_nloc);
            break;
    }
    
    var stype = "";
    if (bwithtype == true) {
        stype = ocurritem.Type();
    }
    
    // AttrList in source language
    var oattrlist = ocurritem.AttrList(g_tcheck.nsourcelang);
    
    if (g_tcheck.bexpand) {
        // BLUE-11810 Expand attribute list by attributes maintained in target language
        oattrlist = expandAttrList(ocurritem, oattrlist, g_tcheck.nsourcelang, g_tcheck.ntargetlang);
    }
    
    for (var i = 0 ; i < oattrlist.length ; i++ ) {
        var ocurrattr = oattrlist[i];
        
        if (ocurrattr.IsChangeable() && ocurrattr.LanguageDependence() != 1) {
            switch(g_omethodfilter.AttrBaseType(ocurrattr.TypeNum())) {
                case Constants.ABT_BITMAP:
                case Constants.ABT_ITEMTYPE:
                case Constants.ABT_LONGTEXT:
                case Constants.ABT_VALUE:
                case Constants.ABT_COMBINED:
                    break;
                default:
                    var nattrtypenum = ocurrattr.TypeNum();
                    var sattrtype = ocurrattr.Type();
                    var ssourceval = getattributevalue(ocurrattr);

                    var olangattr = ocurritem.Attribute(nattrtypenum, g_tcheck.ntargetlang);
                    var stargetval = "";
                    if ((!olangattr.IsMaintained()) || (!g_tcheck.bonlynotmaint)) {
                        if (olangattr.IsMaintained()) {
                            stargetval = getattributevalue(olangattr);
                        }
    
                        g_ooutfile.TableRow();
                        g_ooutfile.TableCell(sguid, g_nwidth, getString("TEXT13"), 10, Constants.C_BLACK, getTableCellColor_AttrBk(g_bColored), 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
                        
                        if (bwithtype) {
                            g_ooutfile.TableCell(stype, g_nwidth, getString("TEXT13"), 10, Constants.C_BLACK, getTableCellColor_AttrBk(g_bColored), 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
                        }
                        g_ooutfile.TableCell(sname, g_nwidth, getString("TEXT13"), 10, Constants.C_BLACK, getTableCellColor_AttrBk(g_bColored), 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
                        g_ooutfile.TableCell(getAttrTypeOrGuid(nattrtypenum), g_nwidth, getString("TEXT13"), 10, Constants.C_BLACK, getTableCellColor_AttrBk(g_bColored), 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
                        g_ooutfile.TableCell(sattrtype, g_nwidth, getString("TEXT13"), 10, Constants.C_BLACK, getTableCellColor_AttrBk(g_bColored), 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
                        g_ooutfile.TableCell(ssourceval, g_nwidth, getString("TEXT13"), 10, Constants.C_BLACK, getTableCellColor_AttrBk(g_bColored), 0, Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
                        
                        if (g_tcheck.bwithtarget) {
                            g_ooutfile.TableCell(stargetval, g_nwidth, getString("TEXT13"), 10, Constants.C_BLACK, getTableCellColor_AttrBk(g_bColored), 0, Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
                        }
                        
                        nrow_holder.value = nrow_holder.value + 1;
                        if (nrow_holder.value == g_nrowmax) {
                            ssheetnametmp = ssheetname + "_" + ncount_holder.value;
                            g_ooutfile.EndTable(ssheetnametmp, 100, getString("TEXT13"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    
                            g_ooutfile.BeginTable(100, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
                            outitemheadlines(getItemKind(ocurritem), bwithtype);
                            g_bColored = false;
                            
                            ncount_holder.value = ncount_holder.value + 1;
                            nrow_holder.value = 2;
                        }
                        g_bColored = !g_bColored;   // Change background color                    
                    }
            }
        }
    }
    
    function expandAttrList(oItem, srcAttrList, srcLang, trgLang) {
        var setSrcAttrTypes = getSrcAttrTypes();
        
        var trgAttrList = oItem.AttrList(trgLang);
        for (var i in trgAttrList) {
            var nAttrTypeNum = trgAttrList[i].TypeNum();
            if (setSrcAttrTypes.contains(nAttrTypeNum)) continue;
            
            srcAttrList.push(oItem.Attribute(nAttrTypeNum, srcLang));   // Add attribute to attribute list
        }
        return srcAttrList;

        function getSrcAttrTypes() {
            var set = new java.util.HashSet();
            for (var i in srcAttrList) {
                set.add(srcAttrList[i].TypeNum());
            }
            return set;
        }
        
    }
}

function outcxnattributes(ocurrcxn, ssheetname, ncount_holder, nrow_holder) {
    var osourceobjdef = ocurrcxn.SourceObjDef();
    var otargetobjdef = ocurrcxn.TargetObjDef();
    
    var sguid = ocurrcxn.GUID();
    var stype = ocurrcxn.ActiveType();
    var ssource = osourceobjdef.Name(g_nloc) + " / " + osourceobjdef.Type();
    var starget = otargetobjdef.Name(g_nloc) + " / " + otargetobjdef.Type();
    
    // AttrList in source language
    var oattrlist = ocurrcxn.AttrList(g_tcheck.nsourcelang);
    for (var i = 0 ; i < oattrlist.length ; i++ ) {
        var ocurrattr = oattrlist[i];
        
        if (ocurrattr.IsChangeable() && ocurrattr.LanguageDependence() != 1) {
            switch(g_omethodfilter.AttrBaseType(ocurrattr.TypeNum())) {
                case Constants.ABT_BITMAP:
                case Constants.ABT_ITEMTYPE:
                case Constants.ABT_LONGTEXT:
                case Constants.ABT_VALUE:
                case Constants.ABT_COMBINED:
                    break;
                default:
                    var nattrtypenum = ocurrattr.TypeNum();
                    var sattrtype = ocurrattr.Type();
                    var ssourceval = getattributevalue(ocurrattr);
                    
                    var olangattr = ocurrcxn.Attribute(nattrtypenum, g_tcheck.ntargetlang);
                    var stargetval = "";
                    if ((!olangattr.IsMaintained()) || (!g_tcheck.bonlynotmaint)) {
                        if (olangattr.IsMaintained()) {
                            stargetval = getattributevalue(olangattr);
                        }
        
                            g_ooutfile.TableRow();
                            g_ooutfile.TableCell(sguid, g_nwidth, getString("TEXT13"), 10, Constants.C_BLACK, getTableCellColor_AttrBk(g_bColored), 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
                            g_ooutfile.TableCell(stype, g_nwidth, getString("TEXT13"), 10, Constants.C_BLACK, getTableCellColor_AttrBk(g_bColored), 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
                            g_ooutfile.TableCell(ssource, g_nwidth, getString("TEXT13"), 10, Constants.C_BLACK, getTableCellColor_AttrBk(g_bColored), 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
                            g_ooutfile.TableCell(starget, g_nwidth, getString("TEXT13"), 10, Constants.C_BLACK, getTableCellColor_AttrBk(g_bColored), 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
                            g_ooutfile.TableCell(getAttrTypeOrGuid(nattrtypenum), g_nwidth, getString("TEXT13"), 10, Constants.C_BLACK, getTableCellColor_AttrBk(g_bColored), 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
                            g_ooutfile.TableCell(sattrtype, g_nwidth, getString("TEXT13"), 10, Constants.C_BLACK, getTableCellColor_AttrBk(g_bColored), 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
                            g_ooutfile.TableCell(ssourceval, g_nwidth, getString("TEXT13"), 10, Constants.C_BLACK, getTableCellColor_AttrBk(g_bColored), 0, Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
                            
                            if (g_tcheck.bwithtarget) {
                                g_ooutfile.TableCell(stargetval, g_nwidth, getString("TEXT13"), 10, Constants.C_BLACK, getTableCellColor_AttrBk(g_bColored), 0, Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
                            }
        
                            nrow_holder.value = nrow_holder.value + 1;
                            if (nrow_holder.value == g_nrowmax) {
                                ssheetnametmp = ssheetname + "_" + ncount_holder.value;
                                g_ooutfile.EndTable(ssheetnametmp, 100, getString("TEXT13"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
        
                                g_ooutfile.BeginTable(100, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
                                outcxnheadlines();
                                g_bColored = false;
                                
                                ncount_holder.value = ncount_holder.value + 1;
                                nrow_holder.value = 2;
                            }
                            g_bColored = !g_bColored;   // Change background color
                    }
            }
        }
    }
}

function getAttrTypeOrGuid(nattrtypenum) {
    // AGA-6929
    if (g_omethodfilter.isUserDefinedAttrType(nattrtypenum)) {
        return g_omethodfilter.UserDefinedAttributeTypeGUID(nattrtypenum);
    }
    return nattrtypenum;
}

function getattributevalue(ocurrattr) {
    return "" + ocurrattr.GetValue(false);
}

function userdlg(tcheck) {
    
    var nuserdlg = 0;   // Variable for the user dialog box
    
    var olanguages = getlanguages();
    var slanguages = new Array(); 
    var nlanguages = new Array();     
    for (var i = 0 ; i < olanguages.length ; i++ ) {
        slanguages.push(getlanguagename(olanguages[i]));
        nlanguages.push(olanguages[i].LocaleId());
    }
    
    var userdialog = Dialogs.createNewDialogTemplate(0, 0, 480, 235, getString("TEXT2"), "dlg_Listener");
    // %GRID:10,7,1,1
    userdialog.Text(40, 15, 100, 14, getString("TEXT32"), "Text1");
    userdialog.DropListBox(150, 10, 280, 70, slanguages, "SourceLang");
    userdialog.Text(40, 45, 90, 14, getString("TEXT33"), "Text2");
    userdialog.DropListBox(150, 45, 280, 70, slanguages, "TargetLang");
    userdialog.GroupBox(40, 77, 400, 110, getString("TEXT34"), "GroupBox1");
    userdialog.CheckBox(60, 100, 170, 15, getString("TEXT12"), "Check_Database");
    userdialog.CheckBox(60, 120, 170, 15, getString("TEXT14"), "Check_Group");
    userdialog.CheckBox(60, 140, 170, 15, getString("TEXT15"), "Check_Model");
    userdialog.CheckBox(60, 160, 170, 15, getString("TEXT16"), "Check_Object");
    userdialog.CheckBox(240, 100, 190, 15, getString("TEXT17"), "Check_Cxn");   
    userdialog.CheckBox(240, 120, 190, 15, getString("TEXT40"), "Check_Shortcut");  
    userdialog.CheckBox(240, 140, 190, 15, getString("TEXT18"), "Check_Text");
    userdialog.CheckBox(240, 160, 190, 15, getString("TEXT20"), "Check_FontStyle");
    userdialog.CheckBox(40, 200, 500, 15, getString("TEXT41"), "Check_Only_NotMaintained");
    userdialog.CheckBox(40, 220, 500, 15, getString("TEXT42"), "Check_Expand");     // BLUE-11810
    userdialog.CheckBox(40, 240, 250, 15, getString("TEXT38"), "Check_Recursion");
    userdialog.OKButton();
    userdialog.CancelButton();
    userdialog.HelpButton("HID_f889d8e0_eaea_11d8_12e0_9d2843560f51_dlg_01.hlp");
    
    var dlg = Dialogs.createUserDialog(userdialog); 
    
    // Read dialog settings from config    
    var sSection = "SCRIPT_f889d8e0_eaea_11d8_12e0_9d2843560f51";
    ReadSettingsListBoxByNumber(dlg, sSection, "SourceLang", 0, nlanguages);
    ReadSettingsListBoxByNumber(dlg, sSection, "TargetLang", 0, nlanguages);
    ReadSettingsDlgValue(dlg, sSection, "Check_Database", 1);
    ReadSettingsDlgValue(dlg, sSection, "Check_Group", 1);
    ReadSettingsDlgValue(dlg, sSection, "Check_Model", 1);
    ReadSettingsDlgValue(dlg, sSection, "Check_Object", 1);
    ReadSettingsDlgValue(dlg, sSection, "Check_Cxn", 1);
    ReadSettingsDlgValue(dlg, sSection, "Check_Shortcut", 1);
    ReadSettingsDlgValue(dlg, sSection, "Check_Text", 1);
    ReadSettingsDlgValue(dlg, sSection, "Check_FontStyle", 1);
    ReadSettingsDlgValue(dlg, sSection, "Check_Only_NotMaintained", 0);
    ReadSettingsDlgValue(dlg, sSection, "Check_Expand", 0);
    ReadSettingsDlgValue(dlg, sSection, "Check_Recursion", 0);
    
    nuserdlg = Dialogs.show( __currentDialog = dlg);            // Show dialog (wait for ok).

    // Write dialog settings to config
    if (nuserdlg != 0) {      
        WriteSettingsListBoxByNumber(dlg, sSection, "SourceLang", nlanguages);
        WriteSettingsListBoxByNumber(dlg, sSection, "TargetLang", nlanguages);
        WriteSettingsDlgValue(dlg, sSection, "Check_Database");
        WriteSettingsDlgValue(dlg, sSection, "Check_Group");
        WriteSettingsDlgValue(dlg, sSection, "Check_Model");
        WriteSettingsDlgValue(dlg, sSection, "Check_Object");
        WriteSettingsDlgValue(dlg, sSection, "Check_Cxn");
        WriteSettingsDlgValue(dlg, sSection, "Check_Shortcut");  
        WriteSettingsDlgValue(dlg, sSection, "Check_Text");
        WriteSettingsDlgValue(dlg, sSection, "Check_FontStyle");
        WriteSettingsDlgValue(dlg, sSection, "Check_Only_NotMaintained");
        WriteSettingsDlgValue(dlg, sSection, "Check_Expand");
        WriteSettingsDlgValue(dlg, sSection, "Check_Recursion");
    }

    tcheck.ndatabase = dlg.getDlgValue("Check_Database");
    tcheck.ngroup = dlg.getDlgValue("Check_Group");
    tcheck.nmodel = dlg.getDlgValue("Check_Model");
    tcheck.nobject = dlg.getDlgValue("Check_Object");
    tcheck.ncxn = dlg.getDlgValue("Check_Cxn");
    tcheck.nshortcut = dlg.getDlgValue("Check_Shortcut");  
    tcheck.ntext = dlg.getDlgValue("Check_Text");
    tcheck.nfontstyle = dlg.getDlgValue("Check_FontStyle");
    tcheck.bonlynotmaint = dlg.getDlgValue("Check_Only_NotMaintained") == 1;
    tcheck.bexpand = dlg.getDlgValue("Check_Expand") == 1;
    tcheck.brecursive = dlg.getDlgValue("Check_Recursion") == 1;
    
    tcheck.nsourcelang = olanguages[dlg.getDlgValue("SourceLang")].LocaleId();
    tcheck.ntargetlang = olanguages[dlg.getDlgValue("TargetLang")].LocaleId();
    
    // only if source language <> target language then output of target language
    tcheck.bwithtarget = false;
    if (! (tcheck.nsourcelang == tcheck.ntargetlang)) {
        tcheck.bwithtarget = true;
    }
    
    if (nuserdlg == 0) {
        return false;
    } 
    return true;
}

//Dialog listener
function dlg_Listener(dlgitem, action, suppvalue){
    var result = false;
    
    switch(action) {
        case 1:
            updateOptions();
            break;
        case 2:
            if (dlgitem == "Check_Only_NotMaintained") {
                updateOptions();
                result = true;
            }
            break;
    }
    return result; 
    
    function updateOptions() {
        if (__currentDialog.getDlgValue("Check_Only_NotMaintained") == 1) {
            __currentDialog.setDlgValue("Check_Expand", 0);         
            __currentDialog.setDlgEnable("Check_Expand", false);
        } else {
            __currentDialog.setDlgEnable("Check_Expand", true);
        }
    }
}

function getlanguages() {
    var olanguages = new Array();

    var olanguagelist = ArisData.getActiveDatabase().LanguageList();   
    // get selected language (-> first entry in list)
    for (var  i = 0 ; i < olanguagelist.length ; i++ ) {
        if (olanguagelist[i].LocaleId() == g_nloc) {
            olanguages.push(olanguagelist[i]);
            break;
        }
    }
    // get all other languages
    for (var i = 0 ; i < olanguagelist.length ; i++ ) {
        if (olanguagelist[i].LocaleId() != g_nloc) {
            olanguages.push(olanguagelist[i]);
        }
    }
    return olanguages;
}

function getlanguagename(olanguage) {
    //  return olanguage.Name(g_nloc); 

    // TANR 241433    
    var oLocaleInfo = olanguage.LocaleInfo()
    var oLocale = oLocaleInfo.getLocale();

    return oLocale.getDisplayName();
}
