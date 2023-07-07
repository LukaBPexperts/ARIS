// static variables
var giLoc = Context.getSelectedLanguage();
var gExcelObject = null;
var giStandardColumnWidth = 15000;
var giRowCounter = 0;
var lDataRowStyle = null;
var gCurrentFilter;
var gsBulletSign = "\u2022";
var gsSheetName = "Printed at ";


const giProcessFlowChartStandard = Constants.MT_EEPC; // Process Flow Chart 
const giProcessFlowChartStandardSwimlanes = Constants.MT_EEPC_COLUMN; // Process Flow Chart - Swimlanes (EPC (column display))


function main() {
    gExcelObject = Context.createExcelWorkbook(Context.getSelectedFile());
    var selectedObjects = ArisData.getSelectedObjDefs();
    if (selectedObjects.length > 0) {
        // save active metod filter
        gCurrentFilter = ArisData.ActiveFilter();
        printRowHeaders(gExcelObject);
        // loop through selected models and output results
        for (var i = 0; i < selectedObjects.length; i++) {
            collectAndPrintData(selectedObjects[i], gExcelObject);
        }
    } else if (selectedModels.length < 1) {
        Dialogs.MsgBox("Invalid selection.\nNo model selection found!");
    }
    // write report
    gExcelObject.write();
}

function printRowHeaders(pExcelWorkBook) {
    var sDate = java.text.SimpleDateFormat("dd-MM-yyyy HH.mm").format(new java.util.Date());
    var sSheetName = gsSheetName + sDate;
    var lMainSheet = pExcelWorkBook.createSheet(sSheetName);
    var lHeaderRow = lMainSheet.createRow(giRowCounter);
    // set column width
    // Part of
    lMainSheet.setColumnWidth(0, giStandardColumnWidth);
    // Activity name
    lMainSheet.setColumnWidth(1, giStandardColumnWidth);
    // Activity Description
    lMainSheet.setColumnWidth(2, giStandardColumnWidth);
    // Input
    lMainSheet.setColumnWidth(3, giStandardColumnWidth);
    // Output
    lMainSheet.setColumnWidth(4, giStandardColumnWidth);


    /* Freeze top row alone */
    lMainSheet.createFreezePane(0, 1);

    // set style
    var lHeaderFont = pExcelWorkBook.createFont();
    lHeaderFont.setBoldweight(Constants.XL_FONT_WEIGHT_BOLD);


    var lCellL1 = lHeaderRow.createCell(0);
    lCellL1.setCellValue("Part of");
    
    //var lCellStyle = lCellL1.getCellStyle();
    //lCellStyle.setFont(lHeaderFont);
    
    var lCellL2 = lHeaderRow.createCell(1);
    lCellL2.setCellValue("Activity name");
    
    var lCellL3 = lHeaderRow.createCell(2);
    lCellL3.setCellValue("Activity Description");
    
    var lCellL4 = lHeaderRow.createCell(3);
    lCellL4.setCellValue("Input");
    
    var lCell5 = lHeaderRow.createCell(4);
    lCell5.setCellValue("Input RASCI");
    
    
    // set new style
    var lDataFont = pExcelWorkBook.createFont();
    lDataFont.setBoldweight(Constants.XL_FONT_WEIGHT_NORMAL);

    var liBorderColor = Constants.C_GRAY;
    lDataRowStyle = pExcelWorkBook.createCellStyle(lDataFont, 1, 1, 1, 1, liBorderColor, liBorderColor, liBorderColor, liBorderColor, Constants.ALIGN_LEFT, Constants.VERTICAL_TOP, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.SOLID_FOREGROUND);
    lDataRowStyle.setWrapText(true);
    // set style
    /* var lCell13 = lHeaderRow.createCell(12);
    lCell13.setCellStyle(lDataRowStyle);
    lCell13.setCellValue(" "); */

    giRowCounter++
}

function collectAndPrintData(pObject, pExcelWorkBook) {
    try {

                // activity counter, used for merge
                var funcBr = 0;
                
        var psObjectName = pObject.Name(giLoc);
        


        //added 04/07/2019
        //row color counter
        var rowColorBr = 1;
        
            printRow(psObjectName, "", "", "", pExcelWorkBook, rowColorBr);
            funcBr++;
        
        /*
        for (var i = 0; i < laObjOcc.length; i++) {
            // Activity (Check for Process interfaces)


                //added 04/07/2019 used for merge columns
                // activity counter
                var funcBr = 0;

                var loActivityObjDef = laObjOcc[i].ObjDef();
                // activity name
                var lsActivityName = laObjOcc[i].ObjDef().Name(giLoc);
                // activity description
                var lsActivityDescription = getItemAttributeValue(loActivityObjDef, 9, "");

                // find role 'Responsible'
                var laResultsResponsible = getRelatedOccurrence(laObjOcc[i], Constants.CT_EXEC_2, Constants.OT_PERS_TYPE, null); // is responsible for	218	CT_EXEC_2 Person type (78)	OT_PERS_TYPE
                lsResponsibleName = ObjsListNamesToString(laResultsResponsible); //added 21/01/2020

                lsTitle1 = getItemAttributeValue(loActivityObjDef, Constants.AT_TITL1, "");
                lsTitle1 = filterCarriageReturn(lsTitle1);

        
                var lmAssignedFAD = getAssignedModel(laObjOcc[i], 14);

           
                    
                        var laResultsScreens = getRelatedOccurrence(loOccurrenceOnFAD[0], Constants.CT_IS_REPR_BY, Constants.OT_SCRN, null); // is informed of	326	CT_MUST_BE_INFO_ABT_2 Person type (78)	OT_PERS_TYPE
                        for each(resultScreen in laResultsScreens) {
                            var lsScreenName = resultScreen.Name(giLoc);
                            lsScreenFullName = getItemAttributeValue(resultScreen, Constants.AT_NAME_FULL, "");
                            printRow(psModelName, lsActivityName, lsActivityDescription, something, rowColorBr);
                            funcBr++;
                        }

                        
                    printRow(psModelName, lsActivityName, lsActivityDescription, something, lsOutputrowColorBr);
                    funcBr++;
                    // save print status Input/output


                //merge cells after print
                //mergeCells(funcBr, pExcelWorkBook) //adjust this
           
        }
        */
        
    } catch (ex) {
        Dialogs.MsgBox("Selection invalid, more then one model selected!" + ex);
    }
}

function getItemAttributeValue(pArisItem, piAttributeTypeNum, sDefaultValue) {
    var sValue = "";
    try {
        var lAttObj = pArisItem.Attribute(piAttributeTypeNum, giLoc);
        if (lAttObj.IsMaintained() == true) {
            sValue = lAttObj.getValue();
        } else {
            sValue = sDefaultValue;
        }
    } catch (ex) {
        Dialogs.MsgBox("Error : " + ex);
    }
    return sValue;
}

function ObjsListNamesToString(paObjList) {
    var psResult = "";
    var liCounter = 0;
    var nrOfElements = paObjList.length;
    for each(loObj in paObjList) {
        if (liCounter + 1 == nrOfElements) {
            psResult += gsBulletSign + " " + loObj.Name(giLoc);
        } else {
            psResult += gsBulletSign + " " + loObj.Name(giLoc) + "\n";
        }
    }
    return psResult;
}

// remove line break from String parameter psText
function filterCarriageReturn(psText) {
    var lsResult = "";
    if (psText != null && psText.equalsIgnoreCase("") == false) {
        lsResult = psText.replaceAll("\\r|\\n", "");
    }
    return lsResult;
}

function mergeCells(funcBr, pExcelWorkBook) {
    //added 04/07/2019
    if (funcBr > 1) {
        var regionSheet = pExcelWorkBook.getSheetAt(0);
        var fromRow = giRowCounter - funcBr
        
        regionSheet.addRegion(0, 0, fromRow, giRowCounter - 1)
        regionSheet.addRegion(1, 1, fromRow, giRowCounter - 1)
        regionSheet.addRegion(2, 2, fromRow, giRowCounter - 1)
        regionSheet.addRegion(8, 8, fromRow, giRowCounter - 1)
        regionSheet.addRegion(9, 9, fromRow, giRowCounter - 1)
        regionSheet.addRegion(10, 10, fromRow, giRowCounter - 1)
        regionSheet.addRegion(11, 11, fromRow, giRowCounter - 1)
        regionSheet.addRegion(12, 12, fromRow, giRowCounter - 1)
        regionSheet.addRegion(13, 13, fromRow, giRowCounter - 1)
        regionSheet.addRegion(14, 14, fromRow, giRowCounter - 1)
        regionSheet.addRegion(15, 15, fromRow, giRowCounter - 1)
        regionSheet.addRegion(16, 16, fromRow, giRowCounter - 1)
        regionSheet.addRegion(17, 17, fromRow, giRowCounter - 1)
        regionSheet.addRegion(18, 18, fromRow, giRowCounter - 1)
        regionSheet.addRegion(19, 19, fromRow, giRowCounter - 1)
        regionSheet.addRegion(20, 20, fromRow, giRowCounter - 1)
        regionSheet.addRegion(21, 21, fromRow, giRowCounter - 1)
    }
}

function printRow(psModelName, psActivityName, psActivityDescription, psInput, pExcelWorkBook, rowColorBr) {
    try {
        
        //added 04/07/2019
        //shade every other row, based on rowColorBr counter
        var crFont = pExcelWorkBook.createFont();
        crFont.setBoldweight(Constants.XL_FONT_WEIGHT_NORMAL);
        var crLiBorderColor = Constants.C_GRAY;
        if (rowColorBr % 2 == 1) {
            crDataRowStyle = pExcelWorkBook.createCellStyle(crFont, 1, 1, 1, 1, crLiBorderColor, crLiBorderColor, crLiBorderColor, crLiBorderColor, Constants.ALIGN_LEFT, Constants.VERTICAL_TOP, Constants.C_TRANSPARENT, Constants.C_GREY_80_PERCENT, Constants.SOLID_FOREGROUND);
        } else {
            crDataRowStyle = pExcelWorkBook.createCellStyle(crFont, 1, 1, 1, 1, crLiBorderColor, crLiBorderColor, crLiBorderColor, crLiBorderColor, Constants.ALIGN_LEFT, Constants.VERTICAL_TOP, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.SOLID_FOREGROUND);
        }
        crDataRowStyle.setWrapText(true);

        var lMainSheet = pExcelWorkBook.getSheetAt(0);

        // second row data
        var lModelRow = lMainSheet.createRow(giRowCounter);

        var lCell0 = lModelRow.createCell(0);
        lCell0.setCellValue(psModelName);
        lCell0.setCellStyle(crDataRowStyle);

        var lCell1 = lModelRow.createCell(1);
        lCell1.setCellValue(psActivityName);
        lCell1.setCellStyle(crDataRowStyle);

        var lCell2 = lModelRow.createCell(2);
        lCell2.setCellValue(psActivityDescription);
        lCell2.setCellStyle(crDataRowStyle);

        // Input Name
        var lCellBPO = lModelRow.createCell(3);
        lCellBPO.setCellValue(psInput);
        lCellBPO.setCellStyle(crDataRowStyle);


        // up counter
        giRowCounter++;
    } catch (ex) {
        Dialogs.MsgBox("Error : " + ex);
    }
}

main()