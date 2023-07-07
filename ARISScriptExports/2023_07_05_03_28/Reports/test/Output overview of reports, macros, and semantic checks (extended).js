/**
 * Copyright (C) 2018 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 */

const COMP_REPORT   = 1;
const COMP_MACRO    = 7;
const COMP_SEMCHECK = 3;

var nLoc = Context.getSelectedLanguage();
var oOut = Context.createOutputObject();

var scriptAdmin = Context.getComponent("ScriptAdmin");

outScripts(COMP_REPORT, getString("REPORTS"));
outScripts(COMP_MACRO, getString("MACROS"));
outScripts(COMP_SEMCHECK, getString("SEMANTICCHECKS"));

oOut.WriteReport();  

/*************************************************************************/  

function outScripts(componentID, componentName) {
    oOut.BeginTable(100, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
    
    if (componentID == COMP_SEMCHECK) {
        outHead();
    }
    else{
        outHeadReports();
    }
    
    var bColored = false;
    
    var aCategories = scriptAdmin.getCategories(componentID, nLoc).sort(sortCategory);
    
    for (var i in aCategories) {
        var category = aCategories[i];
        var aScriptInfos = scriptAdmin.getScriptInfos(componentID, category.getCategoryID(), nLoc).sort(sortName);
        
        if (aScriptInfos.length > 0) {
            var bFirst = true;

            for (var j in aScriptInfos) {
                var scriptInfo = aScriptInfos[j];
                
                if (scriptInfo.isSimpleFile()) continue;
                
                if (componentID == COMP_SEMCHECK) {
                    outScript(getCategoryName(category), scriptInfo, bColored, bFirst);
                }
                else{
                    outScriptReport(getCategoryName(category), scriptInfo, bColored, bFirst,componentID,category.getCategoryID());
                }
                
                bFirst = false;
            }
            
            bColored = !bColored;
        }
    }
    
    oOut.EndTable(componentName, 100, getString("FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0);
    
    function getCategoryName(category) {
        if (componentID == COMP_SEMCHECK) {
            if (category.getCategoryID() == "9b676eb9-ec3d-4b53-b508-a62475f4d433") return getString("RULETYPES");
            if (category.getCategoryID() == "54e057ed-80b1-4521-9545-6496bcd27cd1") return getString("PROFILES");        
        }
        
        return category.getName();
    }
}


function outScript(sCategoryName, scriptInfo, bColored, bFirst) {
    var bgColor = bColored ? RGB( 220, 230, 241) : Constants.C_TRANSPARENT;
    var catColor = bFirst ? Constants.C_BLACK : Constants.C_BLACK;

    oOut.TableRow();
    oOut.TableCell(sCategoryName, 50, getString("FONT"), 10, catColor, bgColor, 0, Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
    oOut.TableCell(scriptInfo.getID(), 40, getString("FONT"), 10, Constants.C_BLACK, bgColor, 0, Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
    oOut.TableCell(scriptInfo.getName(), 50, getString("FONT"), 10, Constants.C_BLACK, bgColor, 0, Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
    oOut.TableCell(scriptInfo.getDescription(), 100, getString("FONT"), 10, Constants.C_BLACK, bgColor, 0, Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
}


function outScriptReport(sCategoryName, scriptInfo, bColored, bFirst,componentID,sCategoryID) {
    var bgColor = bColored ? RGB( 220, 230, 241) : Constants.C_TRANSPARENT;
    var catColor = bFirst ? Constants.C_BLACK : Constants.C_BLACK;
    
    oOut.TableRow();
    oOut.TableCell(scriptInfo.isAvailableForUsers() ? "X" : "" , 15, getString("FONT"), 10, Constants.C_BLACK, bgColor, 0, Constants.FMT_CENTER | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
    oOut.TableCell(scriptInfo.isAllowedInConnect()? "X" : "" , 15, getString("FONT"), 10, Constants.C_BLACK, bgColor, 0, Constants.FMT_CENTER | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
    oOut.TableCell(sCategoryName, 50, getString("FONT"), 10, catColor, bgColor, 0, Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
    oOut.TableCell(scriptInfo.getID(), 40, getString("FONT"), 10, Constants.C_BLACK, bgColor, 0, Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
    oOut.TableCell(scriptInfo.getName(), 50, getString("FONT"), 10, Constants.C_BLACK, bgColor, 0, Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
    oOut.TableCell(scriptInfo.getDescription(), 70, getString("FONT"), 10, Constants.C_BLACK, bgColor, 0, Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
    oOut.TableCell(GetArrayString(getOutputfilemapping(scriptInfo.getSupportedOutputFormats())), 25, getString("FONT"), 10, Constants.C_BLACK, bgColor, 0, Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
    var sImports = getImports(componentID,sCategoryName,scriptInfo,sCategoryID);
    oOut.TableCell(sImports,100, getString("FONT"), 10, Constants.C_BLACK, bgColor, 0, Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
    
}

function outHead() {
    var bgColor = RGB( 54, 96, 146);
    
    oOut.TableRow();
    oOut.TableCell(getString("CATEGORY"), 50, getString("FONT"), 10, Constants.C_WHITE, bgColor, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
    oOut.TableCell(getString("IDENTIFIER"), 40, getString("FONT"), 10, Constants.C_WHITE, bgColor, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
    oOut.TableCell(getString("NAME"), 50, getString("FONT"), 10, Constants.C_WHITE, bgColor, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
    oOut.TableCell(getString("DESCRIPTION"), 100, getString("FONT"), 10, Constants.C_WHITE, bgColor, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
}


function outHeadReports() {
    var bgColor = RGB( 54, 96, 146);

    oOut.TableRow();
    oOut.TableCell("Available for users", 15, getString("FONT"), 10, Constants.C_WHITE, bgColor, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
    oOut.TableCell("Available in Connect", 15, getString("FONT"), 10, Constants.C_WHITE, bgColor, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
    oOut.TableCell(getString("CATEGORY"), 50, getString("FONT"), 10, Constants.C_WHITE, bgColor, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
    oOut.TableCell(getString("IDENTIFIER"), 40, getString("FONT"), 10, Constants.C_WHITE, bgColor, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
    oOut.TableCell(getString("NAME"), 50, getString("FONT"), 10, Constants.C_WHITE, bgColor, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
    oOut.TableCell(getString("DESCRIPTION"), 70, getString("FONT"), 10, Constants.C_WHITE, bgColor, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
    oOut.TableCell("Output formats", 25, getString("FONT"), 10, Constants.C_WHITE, bgColor, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
    oOut.TableCell("Imports", 100, getString("FONT"), 10, Constants.C_WHITE, bgColor, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
    
}


function RGB(r, g, b) {
	return (new java.awt.Color(r/255.0,g/255.0,b/255.0,1)).getRGB();
}

function sortName(a, b) {
    return StrComp(a.getName(), b.getName());
}
function sortCategory(a, b) {
    var result = sortName(a, b);
    
    if (result == 0) result = -1 * StrComp(a.getCategoryID(), b.getCategoryID()); // -1 to sort semchecks in expected order: 1. rule types, 2. profiles
    return result;
}

function GetArrayString(oArray){
    var sArray ="";
    
    for (var t=0; t < oArray.length;t++){
        sArray = sArray +oArray[t];
        
        if (t+1 <oArray.length){
            sArray = sArray + "\n";
        }
    }
    
    return sArray;
}

function getOutputfilemapping(oOutputfiletypes){
    var aOutputs=new Array();
    // loop
    for (var r=0; r<oOutputfiletypes.length;r++){
        if (oOutputfiletypes[r] == Constants.OUT_NOOUTPUT){
            aOutputs.push("No Output");
        }
    }

    for (var r=0; r<oOutputfiletypes.length;r++){
        if (oOutputfiletypes[r] == Constants.OUTPDF){
            aOutputs.push("PDF");
        }
    }  
    
    for (var r=0; r<oOutputfiletypes.length;r++){
        if (oOutputfiletypes[r] == Constants.OUTWORD){
            aOutputs.push("DOC");
        }
    }
    
    for (var r=0; r<oOutputfiletypes.length;r++){
        if (oOutputfiletypes[r] == Constants.OutputDOCX){
            aOutputs.push("DOCX");
        }
    }
    for (var r=0; r<oOutputfiletypes.length;r++){
        if (oOutputfiletypes[r] == Constants.OutputODT){
            aOutputs.push("ODT");
        }
    }
    
    for (var r=0; r<oOutputfiletypes.length;r++){
        if (oOutputfiletypes[r] == Constants.OUTRTF){
            aOutputs.push("RTF");
        }
    }
   
    for (var r=0; r<oOutputfiletypes.length;r++){
        if (oOutputfiletypes[r] == Constants.OUTHTML){
            aOutputs.push("HTML");
        }
    }
    for (var r=0; r<oOutputfiletypes.length;r++){
        if (oOutputfiletypes[r] == Constants.OUTTEXT){
            aOutputs.push("TXT");
        }
    }
    
    for (var r=0; r<oOutputfiletypes.length;r++){
        if (oOutputfiletypes[r] == Constants.OUTEXCEL){
            aOutputs.push("XLS");
        }
    }
    for (var r=0; r<oOutputfiletypes.length;r++){
        if (oOutputfiletypes[r] == Constants.OutputXLSX){
            aOutputs.push("XLSX");
        }
    }
    for (var r=0; r<oOutputfiletypes.length;r++){
        if (oOutputfiletypes[r] == Constants.OUTXML){
            aOutputs.push("XML");
        }
    }
    
    for (var r=0; r<oOutputfiletypes.length;r++){
        if (oOutputfiletypes[r] == Constants.OUT_USERDEFINED){
            aOutputs.push("User defined output format");
        }
    }

    for (var r=0; r<oOutputfiletypes.length;r++){
        if (oOutputfiletypes[r] == Constants.OutputDOCM){
            aOutputs.push("DOCM");
        }
    }
    
    for (var r=0; r<oOutputfiletypes.length;r++){
        if (oOutputfiletypes[r] == Constants.OutputXLSM){
            aOutputs.push("XLSM");
        }
    }

   return aOutputs;
}

function getImports(componentID,sCategoryName,scriptInfo,sCategoryID){
    var aImportedFiles =new Array();
    var aImports = scriptInfo.getImports();
    
    if (aImports.length>0){
        for (var k=0; k <aImports.length;k++){
            var scurImport = aImports[k];
            
            if (isGuid(scurImport)){
                // get report name in category
                var oImportedReport = scriptAdmin.getScriptInfo(componentID, null, scurImport, nLoc ) ;
                
                if (oImportedReport !=null){
                    aImportedFiles.push(oImportedReport.getName() + " ["+scurImport +"]"); 
                }    
            }
            else{
                aImportedFiles.push(scurImport);
            }
        }
    }   
    
    aImportedFiles = aImportedFiles.sort();
    return GetArrayString(aImportedFiles);
}



// Checks, whether it's a GUID
function isGuid(p_sGuid) {
    var sGuid = new String(p_sGuid);
    
    if (sGuid.length != 36) return false;
    if (sGuid.charAt(8)  != '-') return false;
    if (sGuid.charAt(13) != '-') return false;
    if (sGuid.charAt(18) != '-') return false;
    if (sGuid.charAt(23) != '-') return false;
    return true;
}
