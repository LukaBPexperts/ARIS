// static variables

var giLoc = Context.getSelectedLanguage();
var giStandardColumnWidth = 15000;
var giRowCounter = 0;
var lDataRowStyle = null;
var gCurrentFilter;
var gsBulletSign = "\u2022";
var gsSheetName = "Printed at ";

gExcelObject = Context.createExcelWorkbook(Context.getSelectedFile());

printRowHeaders(gExcelObject);

//printRow("ante", "gudelj", "a", "b", gExcelObject)

var selectedModels = ArisData.getSelectedModels();
var selectedModel = selectedModels[0]


printRow("xxx", "000", selectedModel.Name(giLoc), selectedModel.GUID(), selectedModel.Type(), gExcelObject)

var objDefs = selectedModel.ObjDefListByTypes([Constants.OT_STRCT_ELMT])
for(var i=0; i<objDefs.length; i++){
    
    var assignedModelsL1 = objDefs[i].AssignedModels(Constants.MT_EEPC)
    
    for(var j=0; j<assignedModelsL1.length; j++){
        //these are objDefs which can have either next level assigned or last model       
        printRow(selectedModel.Name(giLoc), selectedModel.GUID(), assignedModelsL1[j].Name(giLoc), assignedModelsL1[j].GUID(), assignedModelsL1[j].Type(), gExcelObject)
        var objDefsL2 = assignedModelsL1[j].ObjDefListBySymbols([ArisData.ActiveFilter().UserDefinedSymbolTypeNum("693974af-b0e0-11df-4732-00237d347a8a"), ArisData.ActiveFilter().UserDefinedSymbolTypeNum("ff856510-b087-11df-4732-00237d347a8a")])
        
        for(var k=0; k<objDefsL2.length; k++){
            var assignedModelsL2 = objDefsL2[k].AssignedModels(Constants.MT_EEPC)
            
            for(var l=0; l<assignedModelsL2.length; l++){
                //this if is visak, 
                if(assignedModelsL2[l].TypeNum() == Constants.MT_EEPC_COLUMN){
                 //print data  
                 printRow(assignedModelsL1[j].Name(giLoc), assignedModelsL1[j].GUID(), assignedModelsL2[l].Name(giLoc), assignedModelsL2[l].GUID(), assignedModelsL2[l].Type(), gExcelObject)
                 var a=2
                }
                else if(assignedModelsL2[l].TypeNum() == Constants.MT_EEPC){
                    printRow(assignedModelsL1[j].Name(giLoc), assignedModelsL1[j].GUID(), assignedModelsL2[l].Name(giLoc), assignedModelsL2[l].GUID(), assignedModelsL2[l].Type(), gExcelObject)
                    //get next level
                    var objDefsL3 = assignedModelsL2[l].ObjDefListBySymbols([ArisData.ActiveFilter().UserDefinedSymbolTypeNum("693974af-b0e0-11df-4732-00237d347a8a"), ArisData.ActiveFilter().UserDefinedSymbolTypeNum("ff856510-b087-11df-4732-00237d347a8a")])
                    for(var m=0; m<objDefsL3.length; m++){
                        var assignedModelsL3 = objDefsL3[m].AssignedModels(Constants.MT_EEPC_COLUMN)
                        for(var n=0; n<assignedModelsL3.length; n++){
                            //print data
                            
                            printRow(assignedModelsL2[l].Name(giLoc), assignedModelsL2[l].GUID(), assignedModelsL3[n].Name(giLoc), assignedModelsL3[n].GUID(), assignedModelsL3[n].Type(), gExcelObject)
                            
                            var a=2
                        }
                    }
                }
            }
        }
    }
}

gExcelObject.write();

function printRowHeaders(pExcelWorkBook) {
    var sDate = java.text.SimpleDateFormat("dd-MM-yyyy HH.mm").format(new java.util.Date());
    var sSheetName = gsSheetName + sDate;
    var lMainSheet = pExcelWorkBook.createSheet(sSheetName);
    var lHeaderRow = lMainSheet.createRow(giRowCounter);
    

    lMainSheet.setDefaultColumnWidth(30);

    /* Freeze top row alone */
    lMainSheet.createFreezePane(0, 1);

    // set style
    var lHeaderFont = pExcelWorkBook.createFont();
    lHeaderFont.setBoldweight(Constants.XL_FONT_WEIGHT_BOLD);
    var headLiBorderColor = Constants.C_GRAY;

    headerStyle = pExcelWorkBook.createCellStyle(lHeaderFont, 1, 1, 1, 1, headLiBorderColor, headLiBorderColor, headLiBorderColor, headLiBorderColor, Constants.ALIGN_CENTER, Constants.VERTICAL_CENTER, Constants.C_TRANSPARENT, Constants.C_LIGHT_BLUE, Constants.SOLID_FOREGROUND);

    var lCellL1 = lHeaderRow.createCell(0);
    lCellL1.setCellValue("Parent Name");
    lCellL1.setCellStyle(headerStyle);

    //var lCellStyle = lCellL1.getCellStyle();
    //lCellStyle.setFont(lHeaderFont);

    var lCellL2 = lHeaderRow.createCell(1);
    lCellL2.setCellValue("Parent GUID");
    lCellL2.setCellStyle(headerStyle);

    var lCellL3 = lHeaderRow.createCell(2);
    lCellL3.setCellValue("Model Name");
    lCellL3.setCellStyle(headerStyle);

    var lCellL4 = lHeaderRow.createCell(3);
    lCellL4.setCellValue("Model GUID");
    lCellL4.setCellStyle(headerStyle);
    
    var lCellL5 = lHeaderRow.createCell(4);
    lCellL5.setCellValue("Model Type");
    lCellL5.setCellStyle(headerStyle);

    
    giRowCounter++
}

function printRow(psGParentName, psParentGUID, psModelName, psModelGUID, psModelType, pExcelWorkBook) {
    try {
        
        //added 04/07/2019
        //shade every other row, based on rowColorBr counter
        var crFont = pExcelWorkBook.createFont();
        crFont.setBoldweight(Constants.XL_FONT_WEIGHT_NORMAL);
        var crLiBorderColor = Constants.C_GRAY;

        crDataRowStyle = pExcelWorkBook.createCellStyle(crFont, 1, 1, 1, 1, crLiBorderColor, crLiBorderColor, crLiBorderColor, crLiBorderColor, Constants.ALIGN_LEFT, Constants.VERTICAL_TOP, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.SOLID_FOREGROUND);

        crDataRowStyle.setWrapText(true);

        var lMainSheet = pExcelWorkBook.getSheetAt(0);

        // second row data
        var lModelRow = lMainSheet.createRow(giRowCounter);

        var lCell0 = lModelRow.createCell(0);
        lCell0.setCellValue(psGParentName);
        lCell0.setCellStyle(crDataRowStyle);

        var lCell1 = lModelRow.createCell(1);
        lCell1.setCellValue(psParentGUID);
        lCell1.setCellStyle(crDataRowStyle);

        var lCell2 = lModelRow.createCell(2);
        lCell2.setCellValue(psModelName);
        lCell2.setCellStyle(crDataRowStyle);

        // Input Name
        var lCellBPO = lModelRow.createCell(3);
        lCellBPO.setCellValue(psModelGUID);
        lCellBPO.setCellStyle(crDataRowStyle);
        
        var lCellBPO = lModelRow.createCell(4);
        lCellBPO.setCellValue(psModelType);
        lCellBPO.setCellStyle(crDataRowStyle);


        // up counter
        giRowCounter++;
    } catch (ex) {
        Dialogs.MsgBox("Error : " + ex);
    }
}