/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

// BLUE-17650 - Import/Usage of 'convertertools.js' removed 


// Name of the error file (transfered to the client and shown) 
var g_sErrorFile = "Errors_f86fc130-eaea-11d8-12e0-9d2843560f51.txt"
// Content of the error file
var g_sErrorText = "";

// global declarations
var g_omethodfilter = ArisData.ActiveFilter();
var g_nloc = Context.getSelectedLanguage();

var g_ntargetlang = 0; 

var g_btouch = false;
var itemsToTouchSet = new java.util.HashSet();


function main()
{ 
    var simportfile = new __holder("");
    if (getimportfile(simportfile)) {
        
        ArisData.Save(Constants.SAVE_ONDEMAND);      
        
        if ( userdlg() ) {
            
            ArisData.getActiveDatabase().setAutoTouch(false);     // No touch !!!
            
            var bsheetok = true;
            var bitemok  = true;
            var battrok  = true;
            
            var app = Context.getExcelReader(simportfile.value.getData());
            
            var sheets = app.getSheets();
            for (var i = 0; i < sheets.length; i++) {
                var currentSheet = sheets[i];     
                
                var yposguid = -1;
                var ypostype = -1;
                var yposattr = -1;
                var nKindNum = null;		// AKC-7973
                
                // Find column with attributes
                var ypos = 0;
                
                while (currentSheet.getCell(0, ypos) != null) {          
                    var cell = currentSheet.getCell(1, ypos);
                    if(cell == null) {
                        ypos = ypos + 1;        // Anubis 342158
                        continue;
                    }                    
                    var svalue = cell.getCellValue();
                    
                    if (svalue != "") {
                        if (StrComp(svalue, "GUID") == 0) {
                            yposguid = ypos;
                        }
                        else if (StrComp(svalue, "AttrType") == 0) {
                            ypostype = ypos;
                        }
                        else if (parseInt(svalue) == g_ntargetlang) {
                            yposattr = ypos;
                        }
                        else if (isKindNum(svalue)) {
                            // AKC-7973 Export report (f889d8e0-eaea-11d8-12e0-9d2843560f51) writes kind num to output for more performant search
                            nKindNum = parseInt(svalue);
                        }
                    }
                    ypos++;
                }
                
                if (yposguid >= 0 && ypostype >= 0 && yposattr >= 0) {
                    var xpos = 2;
                    
                    while (currentSheet.getCell(xpos, yposguid) != null) {          
                        
                        var cellGUID = currentSheet.getCell(xpos, yposguid);
                        if(cellGUID == null) {
                            xpos = xpos + 1;        // Anubis 342158
                            continue;
                        }
                        var sguid = cellGUID.getCellValue();            
                        var cellType = currentSheet.getCell(xpos, ypostype);
                        if(cellType == null) {
                            xpos = xpos + 1;        // Anubis 342158
                            continue;
                        }
                        var nattrtypenum = getAttrTypeNum(cellType.getCellValue());       // AGA-6929
                        
                        var cellAttr = currentSheet.getCell(xpos, yposattr);
                        if(cellAttr == null) {
                            xpos = xpos + 1;        // Anubis 342158
                            continue;
                        }
                        var svalue = new java.lang.String(cellAttr.getCellValue());
                        
                        // Breaks in Excel -> Breaks in ARIS
                        svalue = svalue.replaceAll("\r\n", "\n");
                        svalue = svalue.replaceAll("\n", "\r\n");
                        
                        var ocurritem = new __holder(null); 
                        var bitemisvalid = getitembyguid(ocurritem, sguid, nKindNum);
                        if (bitemisvalid) {
                            
                            var ocurrattr = ocurritem.value.Attribute(nattrtypenum, g_ntargetlang);
                            if (ocurrattr.IsValid()) {
                                
                                var soldValue = ocurrattr.getValue();   
                                if (StrComp(soldValue, svalue) != 0) {      // Write only changed values (TANR 240279)
                                    
                                    changeAutoTouch(ocurritem.value);       // Change AutoTouch (TANR 240279)
                                    
                                    if (ocurrattr.setValue(svalue) == false) {
                                        var smsg = getString("TEXT1") + ocurritem.value.Name(g_nloc) + getString("TEXT2") + ocurrattr.Type() + getString("TEXT3");
                                        writeMsg(smsg);
                                        
                                        battrok = false;
                                    }
                                    
                                    ArisData.getActiveDatabase().setAutoTouch(false);     // No touch !!!
                                }
                            } else {
                                var smsg = getString("TEXT1") + ocurritem.value.Name(g_nloc) + getString("TEXT4") + nattrtypenum;
                                writeMsg(smsg);
                                
                                battrok = false;
                            }
                        } else {
                            var smsg = getString("TEXT5") + sguid + getString("TEXT6");
                            writeMsg(smsg);                
                            
                            bitemok = false;
                        }
                        
                        if (xpos % 100 == 0) {                
                            ArisData.Save(Constants.SAVE_NOW);    // Store every 100 rows per sheet (AKC-7973)
                        }
                        
                        xpos++;
                    }
                    ArisData.Save(Constants.SAVE_NOW);
                } else {
                    if (currentSheet.getCell(0,0) != null) {
                        if (yposattr == -1) {
                            
                            var smsg = getString("TEXT7") + app.getSheetName(i) + getString("TEXT8");
                            writeMsg(smsg);
                        }
                        bsheetok = false;
                    }
                }
            }
            
            if (battrok == false) {
                Dialogs.MsgBox(getString("TEXT9"), Constants.MSGBOX_BTN_OK, getString("TEXT10"));
            }
            
            if (bitemok == false) {
                Dialogs.MsgBox(getString("TEXT11"), Constants.MSGBOX_BTN_OK, getString("TEXT10"));
            }
            
            if (bsheetok == false) {
                Dialogs.MsgBox(getString("TEXT12"), Constants.MSGBOX_BTN_OK, getString("TEXT10"));
            }
            
            if (g_sErrorText.length > 0) {
                var oOut = Context.createOutputObject(Constants.OUTTEXT, g_sErrorFile);
                oOut.OutputLn(g_sErrorText, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
                oOut.WriteReport();
                Dialogs.shell(g_sErrorFile);
                Context.deleteFile(g_sErrorFile);
            }
            Context.setScriptError(Constants.ERR_NOFILECREATED);
        } else {
            Context.setScriptError(Constants.ERR_CANCEL);
        }
        
        ArisData.Save(Constants.SAVE_IMMEDIATELY);          // Ensure that everything is stored (AKC-7973)
        ArisData.getActiveDatabase().setAutoTouch(true);
    } else {
        Dialogs.MsgBox(getString("TEXT14"), Constants.MSGBOX_BTN_OK, getString("TEXT10"));
        Context.setScriptError(Constants.ERR_CANCEL);
    }
}

function isKindNum(sValue) {
    // AKC-7973
    if (isNaN(sValue)) return false;

    var nKindNum = parseInt(sValue);
    switch (nKindNum) {
        case Constants.CID_GROUP:
        case Constants.CID_MODEL:
        case Constants.CID_OBJDEF:
        case Constants.CID_CXNDEF:
            return true;    
    }
    return false;
}

function  getAttrTypeNum(sAttrTypeOrGuid) {
    // AGA-6929
    if (isNaN(sAttrTypeOrGuid)) {
        if (ATSALL.isGuid(sAttrTypeOrGuid)) {
            // userdefined attribute type number
            return g_omethodfilter.UserDefinedAttributeTypeNum(sAttrTypeOrGuid);  
        } else {
            return -1;
        }
    }
    return parseInt(sAttrTypeOrGuid);
}

function changeAutoTouch(oitem) {
    if (g_btouch) {
        if ((oitem.KindNum() == Constants.CID_OBJDEF || oitem.KindNum() == Constants.CID_MODEL || oitem.KindNum() == Constants.CID_TEXTDEF)) {

            if (itemsToTouchSet.add(oitem.GUID())) {
                ArisData.getActiveDatabase().setAutoTouch(true);
            }
        }
    }
}

function getitembyguid(ocurritem, sguid, nKindNum)
{
    // Init
    ocurritem.value = null;
    var bitemisvalid = false;
    
    if (sguid.startsWith("Database_")) {                 // sGuid = "Database_" + <Database.Name>
        var sentry = sguid.substring("Database_".length);
        
        if (StrComp(ArisData.getActiveDatabase().Name(g_nloc), sentry) == 0) {
            ocurritem.value = ArisData.getActiveDatabase();
            bitemisvalid = true;
        }
    } else if (sguid.startsWith("LocaleID_")) {          // sGuid = "LocaleID_" + <Language.localeId>
        var sentry = sguid.substring("LocaleID_".length);
        
        var olanguagelist = ArisData.getActiveDatabase().LanguageList();
        if (olanguagelist.length > 0) {
            for (var i = 0; i < olanguagelist.length; i++) {
                var ocurrlanguage = olanguagelist[i];
                
                if (StrComp(ocurrlanguage.localeId(), sentry) == 0) {
                    ocurritem.value = ocurrlanguage;
                    bitemisvalid = true;
                }
                
                if (bitemisvalid) {
                    break;
                }
            }
        }
    } else {                                                            // sGuid = <GUID>
        if (nKindNum != null) {
            ocurritem.value = ArisData.getActiveDatabase().FindGUID(sguid, nKindNum);     // AKC-7973 More performent search
        } else {
            ocurritem.value = ArisData.getActiveDatabase().FindGUID(sguid);
        }
        bitemisvalid = ocurritem.value.IsValid();
    }
    return bitemisvalid;
}


function userdlg()
{
    var slanguages = new Array();
    
    var olanguages = getlanguages();
    for (var i = 0; i < olanguages.length; i++) {
        slanguages[i] = getlanguagename(olanguages[i]);
    }
    
    var userdialog = Dialogs.createNewDialogTemplate(0, 0, 400, 165, getString("TEXT10"));
    // %GRID:10,7,1,1
    userdialog.GroupBox(30, 15, 340, 60, getString("TEXT15"), "GroupBox1");
    userdialog.DropListBox(60, 40, 280, 70, slanguages, "TargetLang");
    userdialog.CheckBox(30, 85, 20, 14, "", "CheckTouch");
    userdialog.Text(60, 85, 310, 28, getString("TEXT17"));  
    userdialog.OKButton();
    userdialog.CancelButton();
    //  userdialog.HelpButton("HID_f86fc130_eaea_11d8_12e0_9d2843560f51_dlg_01.hlp");
    
    var dlg = Dialogs.createUserDialog(userdialog); 
    
    // Read dialog settings from config    
    var sSection = "SCRIPT_f86fc130_eaea_11d8_12e0_9d2843560f51";
    ReadSettingsDlgValue(dlg, sSection, "TargetLang", 0);
    ReadSettingsDlgValue(dlg, sSection, "CheckTouch", 0);  
    
    var nuserdlg = Dialogs.show( __currentDialog = dlg);		// Show dialog (wait for ok).
    
    // Write dialog settings to config
    if (nuserdlg != 0) {
        WriteSettingsDlgValue(dlg, sSection, "TargetLang");
        WriteSettingsDlgValue(dlg, sSection, "CheckTouch");
    }
    
    g_ntargetlang = olanguages[dlg.getDlgValue("TargetLang")].LocaleId();
    
    g_btouch = dlg.getDlgValue("CheckTouch") == 1;
    
    return nuserdlg != 0;
}


function getlanguages()
{
    var olanguages = new Array();
    var olanguagelist = ArisData.getActiveDatabase().LanguageList();
    
    // Selected language
    for (var i = 0; i < olanguagelist.length; i++) {
        if (olanguagelist[i].LocaleId() == g_nloc) {
            olanguages[olanguages.length] = olanguagelist[i];
            break;
        }
    }
    
    // all other languages
    for (var i = 0 ; i < olanguagelist.length; i++) {
        if (olanguagelist[i].LocaleId() != g_nloc) {
            olanguages[olanguages.length] = olanguagelist[i];
        }
    }
    return olanguages;
}

function getlanguagename(olanguage) {
    // TANR 241433    
    var oLocaleInfo = olanguage.LocaleInfo()
    var oLocale = oLocaleInfo.getLocale();

    return oLocale.getDisplayName();
}

function getimportfile(sImportFile)
{
    // Init
    var sdefname = "";
    var sdefext  = "*.xls!!Excel|*.xls; *.xlsx||";     // BLUE-10385 Support XLSX format
    var sdefdir  = Context.getProfileString("Report", "Output Directory", "");
    var stitle   = getString("TEXT16");
    
    var sfiles = Dialogs.BrowseForFiles(sdefname, sdefext, sdefdir, stitle, 0);
    
    if (sfiles != null && sfiles.length > 0) {
        sImportFile.value = sfiles[0];
        
        return true;
    }
    return false;
}

function writeMsg(smsg) {
    Context.writeOutput(smsg);
    
    if (g_sErrorText.length > 0) g_sErrorText += "\r\n";
    g_sErrorText += smsg;
}


main();