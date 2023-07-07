var giLoc = Context.getSelectedLanguage();
var gExcelObject = null;
var giStandardColumnWidth=15000;
var giRowCounter = 0;
var lDataRowStyle = null;
var gCurrentFilter;
var gsFileLocation="D:\\ARIS_Report_Output\\";
var gsFileName="Vestas VPP extract for GIT ServiceNow.xlsx";

var gaModelsL4Process;

var ghmModelsL4Process = new java.util.HashMap();

var ghsModelsPassed = new java.util.HashSet();

const giProcessFlowChartStandard = Constants.MT_EEPC; // Process Flow Chart 
const giProcessFlowChartStandardSwimlanes = Constants.MT_EEPC_COLUMN; // Process Flow Chart - Swimlanes (EPC (column display))
const giStatusReleaseAttribute =  ArisData.ActiveFilter().UserDefinedAttributeTypeNum("adde7a20-75f9-11e0-61d8-00237d347a8a"); // status
const gsApprovedReleasedStatus = "Released";

// Constants
const gsBODelegateAttrGUID = "618d4120-ea3f-11de-4732-00237d347a8a";
const gsAttVersionGUID = "db3b0fc0-5fd4-11de-4732-00237d347a8a";

function main(){
    gExcelObject = Context.createExcelWorkbook(Context.getSelectedFile()); 
    var selectedModels = ArisData.getSelectedModels();
    if(ArisData.getSelectedDatabases().length==1){
        // save active metod filter
        gCurrentFilter = ArisData.ActiveFilter();
        printRowHeaders(gExcelObject);
        var modelsFound = getProcessModels();
        if (modelsFound==true){
            for each(model in gaModelsL4Process){
                printData(model, gExcelObject, 4);
            }
        }
    }
    else if (selectedModels.length > 1){
        Dialogs.MsgBox("Selection invalid, more then one model selected!");
    }
    // write report
    gExcelObject.write();
    var xlsTemplateData = Context.getFile(Context.getSelectedFile(), Constants.LOCATION_OUTPUT);
    writeContentToSpecificFolder(gsFileLocation,gsFileName,xlsTemplateData);
}
/*
* write byte array to specific file location
*/
function writeContentToSpecificFolder(psFolderPath, psFileName, paFileContent){
    var file = new java.io.File(psFolderPath + psFileName);         
    var fos = null;
    try{
        fos = new java.io.FileOutputStream(file);
        fos.write(paFileContent);
    }catch(e){
        Dialogs.MsgBox("Exception during writing of file: " + e);
    }
    try {
        if (fos != null) {
            fos.close();
        }
    }
    catch (ioe) {
        Dialogs.MsgBox("Error while closing stream: " + ioe);
    }    
}
function getProcessModels(){
    try{
        gaModelsL4Process = ArisData.getActiveDatabase().Find(Constants.SEARCH_MODEL, [giProcessFlowChartStandardSwimlanes]);
        for each(model in gaModelsL4Process){
            ghmModelsL4Process.put(model.GUID(),model);
        }
        return true;
    }
    catch (ex){
        Dialogs.MsgBox("Error : " + ex );
        return false;
    }
}

function printRowHeaders(pExcelWorkBook){
    var lMainSheet = pExcelWorkBook.createSheet("Output");
    var lHeaderRow = lMainSheet.createRow(giRowCounter);
    var liGUIDColumnWidth = 10000;
    var liShortStringColumnWidth = 2500;
    var liBorderColor = Constants.C_GRAY;
    
    // set column width
    // Unique Identifier
    lMainSheet.setColumnWidth (0, liGUIDColumnWidth);
    // Name
    lMainSheet.setColumnWidth (1, giStandardColumnWidth);
    // Description
    lMainSheet.setColumnWidth (2, giStandardColumnWidth);
    // Responsible
    lMainSheet.setColumnWidth (3, liGUIDColumnWidth - liShortStringColumnWidth);
    // Owned by
    lMainSheet.setColumnWidth (4, giStandardColumnWidth);
    // Version
    lMainSheet.setColumnWidth (5, liShortStringColumnWidth);
    
    
    /* Freeze top row alone */
    lMainSheet.createFreezePane(0,1);
    
    // create styles and fonts
    
    var lDataFont = pExcelWorkBook.createFont();
    lDataFont.setBoldweight(Constants.XL_FONT_WEIGHT_NORMAL);
    
    var lHeaderFont = pExcelWorkBook.createFont();
    lHeaderFont.setBoldweight(Constants.XL_FONT_WEIGHT_BOLD); 
    lHeaderFont.setColor(java.awt.Color.WHITE);
    lHeaderFont.setFontHeight(280);
    
    var lCellStyleHeaderGreen = pExcelWorkBook.createCellStyle(lHeaderFont, 1, 1, 1, 1, liBorderColor, liBorderColor, liBorderColor, liBorderColor, Constants.ALIGN_LEFT, Constants.VERTICAL_TOP, Constants.C_TRANSPARENT,Constants.C_TRANSPARENT,Constants.SOLID_FOREGROUND);
    var lCellStyleHeaderBlue = pExcelWorkBook.createCellStyle(lHeaderFont, 1, 1, 1, 1, liBorderColor, liBorderColor, liBorderColor, liBorderColor, Constants.ALIGN_LEFT, Constants.VERTICAL_TOP, Constants.C_TRANSPARENT,getColorByRGB(8,153,204),Constants.SOLID_FOREGROUND);
          
    
    lCellStyleHeaderGreen.setFont(lHeaderFont);
    lCellStyleHeaderGreen.setFillForegroundColor(getColorByRGB(169,208,142));
    lCellStyleHeaderGreen.setFillPattern(Constants.SOLID_FOREGROUND);
    
    // set style
    var lCellL1 = lHeaderRow.createCell(0);
    lCellL1.setCellValue("Unique Identifier");
    lCellL1.setCellStyle(lCellStyleHeaderGreen)
    
    var lCellL2 = lHeaderRow.createCell(1);
    lCellL2.setCellValue("Name");
    lCellL2.setCellStyle(lCellStyleHeaderBlue);    
    var lCellL3 = lHeaderRow.createCell(2);
    lCellL3.setCellValue("Description");
    lCellL3.setCellStyle(lCellStyleHeaderBlue);
    var lCellL4 = lHeaderRow.createCell(3);
    lCellL4.setCellValue("Responsible");
    lCellL4.setCellStyle(lCellStyleHeaderBlue);
    
    var lCell5 = lHeaderRow.createCell(4);
    lCell5.setCellValue("Owned by");
    lCell5.setCellStyle(lCellStyleHeaderGreen)
    
    var lCell6 = lHeaderRow.createCell(5);
    lCell6.setCellValue("Version");
    lCell6.setCellStyle(lCellStyleHeaderGreen)    

    // set style
    lDataRowStyle = pExcelWorkBook.createCellStyle(lDataFont, 1, 1, 1, 1, liBorderColor, liBorderColor, liBorderColor, liBorderColor, Constants.ALIGN_LEFT, Constants.VERTICAL_TOP, Constants.C_TRANSPARENT,Constants.C_TRANSPARENT,Constants.SOLID_FOREGROUND);
    lDataRowStyle.setWrapText(true);
    var lCell13 = lHeaderRow.createCell(6);
    lCell13.setCellStyle(lDataRowStyle);
    lCell13.setCellValue(" ");
    
    giRowCounter++
} 
function getColorByRGB(R, G, B) {
    return (new java.awt.Color(R/255.0 ,G/255.0, B/255.0, 1)).getRGB() & 0xFFFFFF;
} 
//  checks if model attribute 'Model Ready for VPP Publish' has value :'Releaed
//  returns 'true', otherwise 'false' 
function modelIsReleased(pModel){
    var lbResultIsModelReleased=false;
    try{
        var lReleasedAttribute = pModel.Attribute(giStatusReleaseAttribute,giLoc);
        if (lReleasedAttribute.IsMaintained()==true 
        && lReleasedAttribute.getValue().equalsIgnoreCase(gsApprovedReleasedStatus)){
            lbResultIsModelReleased=true;
        }
    }catch(ex){        
        Dialogs.MsgBox("Error while checking if model attribute 'Status' contains value 'Released'." + ex);
    }
    return lbResultIsModelReleased;
}
function printData(pModel, pExcelWorkBook, piModelLevel){
    try{
        // don't process models that are don't have the attribute status 'Released'
        if (modelIsReleased(pModel)==true){
            if (ghsModelsPassed.contains("" + pModel.GUID())==false){
                
                var sModelGUID = pModel.GUID();
                ghsModelsPassed.add("" + sModelGUID);
                // Version
                var liVersionAttTypeNum = ArisData.ActiveFilter().UserDefinedAttributeTypeNum(gsAttVersionGUID);                
                // BO Delegate
                var liBODelegateAttTypeNum = ArisData.ActiveFilter().UserDefinedAttributeTypeNum(gsBODelegateAttrGUID);
                
                var sName = pModel.Name(giLoc);
    
                var lsDescription = getItemAttributeValue(pModel, 9, "");
                var lsResponsible = getItemAttributeValue(pModel,  Constants.AT_PERS_RESP, "");
                
                var lsVersion = getItemAttributeValue(pModel, liVersionAttTypeNum, "");                
                var lsBODelegate = getItemAttributeValue(pModel, liBODelegateAttTypeNum,"");
                                
                printSheetRowModel(sModelGUID, sName, lsDescription, lsResponsible, lsBODelegate, lsVersion, pExcelWorkBook);                
            }
            else{
                var sValue = "";
            }
        }else{
            // model passed because it has not the 'Status','Released'.   
        }
    }
    catch(ex){
        Dialogs.MsgBox("Selection invalid, more then one model selected!" + ex);
    }
}
function getItemAttributeValue(pArisItem, piAttributeTypeNum, sDefaultValue){    
    var sValue = "";
    try{
        var lAttObj = pArisItem.Attribute(piAttributeTypeNum, giLoc);
        if (lAttObj.IsMaintained()==true){
            sValue = lAttObj.getValue();
        }else{
            sValue = sDefaultValue;
        } 
    }
    catch(ex){
        Dialogs.MsgBox("Error : " + ex );
    }
    return sValue;
}
function printSheetRowModel(psModelGUID, psModelName, psDescription, psResponsible, psBODelegate, psVersion, pExcelWorkBook){ 
    try{
        var liCellCounter=0;
        var lMainSheet = pExcelWorkBook.getSheetAt(0);
        
        // second row data
        var lModelRow = lMainSheet.createRow(giRowCounter);
        lDataRowStyle.setWrapText(true);
        
        // Model GUID
        var lCell0 = lModelRow.createCell(liCellCounter);
        liCellCounter++;
        lCell0.setCellValue(psModelGUID);
        lDataRowStyle.setWrapText(true);
        lCell0.setCellStyle(lDataRowStyle);
        
        // Model Name
        var lCell1 = lModelRow.createCell(liCellCounter);
        liCellCounter++;
        lDataRowStyle.setWrapText(true);
        lCell1.setCellValue(psModelName);
        lCell1.setCellStyle(lDataRowStyle);
        
        // Description
        var lCell2 = lModelRow.createCell(liCellCounter);
        liCellCounter++;
        lDataRowStyle.setWrapText(true);
        lCell2.setCellValue(psDescription);
        lCell2.setCellStyle(lDataRowStyle);
        
        // Responsible
        var lCell3 = lModelRow.createCell(liCellCounter);
        liCellCounter++;
        lCell3.setCellValue(psResponsible);
        lDataRowStyle.setWrapText(true);
        lCell3.setCellStyle(lDataRowStyle);
                
        // Owned by
        var lCellOwnedBy = lModelRow.createCell(liCellCounter);
        liCellCounter++;
        lCellOwnedBy.setCellValue(psBODelegate);
        lDataRowStyle.setWrapText(true);
        lCellOwnedBy.setCellStyle(lDataRowStyle);
        
        // Version
        var lCellDesc = lModelRow.createCell(liCellCounter);
        liCellCounter++;
        lCellDesc.setCellValue("" + psVersion.toString() +" ");
        lDataRowStyle.setWrapText(true);
        lCellDesc.setCellStyle(lDataRowStyle);
                        
        // up counter
        giRowCounter++;
    }
    catch(ex){
        Dialogs.MsgBox("Error : " + ex );
    }
}

// remove line break from String parameter psText
function filterCarriageReturn(psText){
    var lsResult = "";
    lsResult = psText.replaceAll("\\r|\\n", "");
    return lsResult;
}

main()