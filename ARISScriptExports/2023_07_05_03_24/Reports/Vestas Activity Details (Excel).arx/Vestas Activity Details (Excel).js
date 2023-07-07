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

/*  
 *   Desc: Main method
 */
function main() {
    gExcelObject = Context.createExcelWorkbook(Context.getSelectedFile());
    var selectedModels = ArisData.getSelectedModels();
    if (selectedModels.length > 0) {
        // save active metod filter
        gCurrentFilter = ArisData.ActiveFilter();
        printRowHeaders(gExcelObject);
        // loop through selected models and output results
        for (var i = 0; i < selectedModels.length; i++) {
            collectAndPrintData(selectedModels[i], gExcelObject);
        }
    } else if (selectedModels.length < 1) {
        Dialogs.MsgBox("Invalid selection.\nNo model selection found!");
    }
    // write report
    gExcelObject.write();
}

/*  
 *   Desc: creates first Row that is used as a column header/label for results
 */
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
    // Output description
    lMainSheet.setColumnWidth(5, giStandardColumnWidth);
    // Supporting Systems
    lMainSheet.setColumnWidth(6, giStandardColumnWidth);
    // R - Responsible
    lMainSheet.setColumnWidth(7, giStandardColumnWidth);
    // A - Accountable
    lMainSheet.setColumnWidth(8, giStandardColumnWidth);
    // S - Supported by
    lMainSheet.setColumnWidth(9, giStandardColumnWidth);
    // C - Consulted for
    lMainSheet.setColumnWidth(10, giStandardColumnWidth);
    // I - Informed of
    lMainSheet.setColumnWidth(11, giStandardColumnWidth);
    // Title 1
    lMainSheet.setColumnWidth(12, giStandardColumnWidth);
    // Link 1
    lMainSheet.setColumnWidth(13, giStandardColumnWidth);
    // Title 2
    lMainSheet.setColumnWidth(14, giStandardColumnWidth);
    // Link 2
    lMainSheet.setColumnWidth(15, giStandardColumnWidth);
    // Title 3
    lMainSheet.setColumnWidth(16, giStandardColumnWidth);
    // Link 3
    lMainSheet.setColumnWidth(17, giStandardColumnWidth);
    // Title 4
    lMainSheet.setColumnWidth(18, giStandardColumnWidth);
    // Link 4
    lMainSheet.setColumnWidth(19, giStandardColumnWidth);
    // SAP T-Codes
    lMainSheet.setColumnWidth(20, giStandardColumnWidth);
    // SAP T-Codes description
    lMainSheet.setColumnWidth(21, giStandardColumnWidth);

    /* Freeze top row alone */
    lMainSheet.createFreezePane(0, 1);

    // set style
    var lHeaderFont = pExcelWorkBook.createFont();
    lHeaderFont.setBoldweight(Constants.XL_FONT_WEIGHT_BOLD);


    var lCellL1 = lHeaderRow.createCell(0);
    lCellL1.setCellValue("Part of");
    var lCellStyle = lCellL1.getCellStyle();
    lCellStyle.setFont(lHeaderFont);
    var lCellL2 = lHeaderRow.createCell(1);
    lCellL2.setCellValue("Activity name");
    var lCellL3 = lHeaderRow.createCell(2);
    lCellL3.setCellValue("Activity Description");
    var lCellL4 = lHeaderRow.createCell(3);
    lCellL4.setCellValue("Input");
    var lCell5 = lHeaderRow.createCell(4);
    lCell5.setCellValue("Output");
    var lCell6 = lHeaderRow.createCell(5);
    lCell6.setCellValue("Input - Output description");
    var lCell7 = lHeaderRow.createCell(6);
    lCell7.setCellValue("Supporting Systems");
    var lCell8 = lHeaderRow.createCell(7);
    lCell8.setCellValue("R - Responsible");
    var lCell9 = lHeaderRow.createCell(8);
    lCell9.setCellValue("A - Accountable");
    var lCell10 = lHeaderRow.createCell(9);
    lCell10.setCellValue("S - Supported by");
    var lCell11 = lHeaderRow.createCell(10);
    lCell11.setCellValue("C - Consulted for");
    var lCell12 = lHeaderRow.createCell(11);
    lCell12.setCellValue("I - Informed of");

    var lCell13 = lHeaderRow.createCell(12);
    lCell13.setCellValue("Title1");
    var lCell14 = lHeaderRow.createCell(13);
    lCell14.setCellValue("Link1");
    var lCell15 = lHeaderRow.createCell(14);
    lCell15.setCellValue("Title2");
    var lCell16 = lHeaderRow.createCell(15);
    lCell16.setCellValue("Link2");
    var lCell17 = lHeaderRow.createCell(16);
    lCell17.setCellValue("Title3");
    var lCell18 = lHeaderRow.createCell(17);
    lCell18.setCellValue("Link3");
    var lCell19 = lHeaderRow.createCell(18);
    lCell19.setCellValue("Title4");
    var lCell20 = lHeaderRow.createCell(19);
    lCell20.setCellValue("Link4");
    //changed 08/07/2019
    var lCell21 = lHeaderRow.createCell(20);
    lCell21.setCellValue("SAP T-Codes Full name");
    var lCell22 = lHeaderRow.createCell(21);
    lCell22.setCellValue("SAP T-Codes");

    // set new style
    var lDataFont = pExcelWorkBook.createFont();
    lDataFont.setBoldweight(Constants.XL_FONT_WEIGHT_NORMAL);

    var liBorderColor = Constants.C_GRAY;
    lDataRowStyle = pExcelWorkBook.createCellStyle(lDataFont, 1, 1, 1, 1, liBorderColor, liBorderColor, liBorderColor, liBorderColor, Constants.ALIGN_LEFT, Constants.VERTICAL_TOP, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.SOLID_FOREGROUND);
    lDataRowStyle.setWrapText(true);
    // set style
    var lCell13 = lHeaderRow.createCell(12);
    lCell13.setCellStyle(lDataRowStyle);
    lCell13.setCellValue(" ");

    giRowCounter++
}
/*  
 *   Desc: finds assigned model that is of model type number : piModelTypeNum
 */
function collectAndPrintData(pModel, pExcelWorkBook) {
    try {

        var laObjOcc = pModel.ObjOccList();
        var psModelName = pModel.Name(giLoc);

        //added 04/07/2019
        //row color counter
        var rowColorBr = 1;

        //added 04/07/2019
        //sort objects by name
        laObjOcc = ArisData.sort(laObjOcc, Constants.AT_NAME, Context.getSelectedLanguage()) //

        for (var i = 0; i < laObjOcc.length; i++) {
            // Activity (Check for Process interfaces)
            if (laObjOcc[i].ObjDef().TypeNum().toString().equalsIgnoreCase("22") &&
                isValidActivity(laObjOcc[i]) == true) {

                var lsInputOutputFound = false;
                var lsInput = "";
                var lsOutput = "";
                var lsInputOutputDescription = "";
                var lsApplicationSystemTypeName = "";
                var lsResponsibleName = "";
                var lsAccountableName = "";
                var lsConsultedName = "";
                var lsSupportedName = "";
                var lsInformedName = "";
                var lsTitle1 = "";
                var lsTitle2 = "";
                var lsTitle3 = "";
                var lsTitle4 = "";

                var lsLink1 = "";
                var lsLink2 = "";
                var lsLink3 = "";
                var lsLink4 = "";

                var lsSAPNames = "";
                var lsSAPFullNames = "";

                //added 04/07/2019
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
                
                // find role 'Accountable'
                var laResultsAccountable = getRelatedOccurrence(laObjOcc[i], Constants.CT_IS_TECH_RESP_3, Constants.OT_PERS_TYPE, null); // is accountable for	220	CT_IS_TECH_RESP_3 Person type (78)	OT_PERS_TYPE
                lsAccountableName = ObjsListNamesToString(laResultsAccountable); //added 21/01/2020
                
                // find role 'Supported'
                var laResultsSupported = getRelatedOccurrence(laObjOcc[i], Constants.CT_DECD_ON, Constants.OT_PERS_TYPE, null); // is supporting	323	CT_DECD_ON Person type (78)	OT_PERS_TYPE
                lsSupportedName = ObjsListNamesToString(laResultsSupported); //added 21/01/2020
                
                // find role 'Consulted'
                var laResultsConsulted = getRelatedOccurrence(laObjOcc[i], Constants.CT_HAS_CONSLT_ROLE_IN_2, Constants.OT_PERS_TYPE, null); // is consulted for	358	CT_HAS_CONSLT_ROLE_IN_2 Person type (78)	OT_PERS_TYPE
                lsConsultedName = ObjsListNamesToString(laResultsConsulted); //added 21/01/2020
                
                // find role 'Informed'
                var laResultsInformed = getRelatedOccurrence(laObjOcc[i], Constants.CT_MUST_BE_INFO_ABT_2, Constants.OT_PERS_TYPE, null); // is informed of	326	CT_MUST_BE_INFO_ABT_2 Person type (78)	OT_PERS_TYPE
                lsInformedName = ObjsListNamesToString(laResultsInformed); //added 21/01/2020
                
                lsTitle1 = getItemAttributeValue(loActivityObjDef, Constants.AT_TITL1, "");
                lsTitle1 = filterCarriageReturn(lsTitle1);

                lsTitle2 = getItemAttributeValue(loActivityObjDef, Constants.AT_TITL2, "");
                lsTitle2 = filterCarriageReturn(lsTitle2);

                lsTitle3 = getItemAttributeValue(loActivityObjDef, Constants.AT_TITL3, "");
                lsTitle3 = filterCarriageReturn(lsTitle3);

                lsTitle4 = getItemAttributeValue(loActivityObjDef, Constants.AT_TITL4, "");
                lsTitle4 = filterCarriageReturn(lsTitle4);

                lsLink1 = getItemAttributeValue(loActivityObjDef, Constants.AT_EXT_1, "");
                lsLink2 = getItemAttributeValue(loActivityObjDef, Constants.AT_EXT_2, "");
                lsLink3 = getItemAttributeValue(loActivityObjDef, Constants.AT_EXT_3, "");
                lsLink4 = getItemAttributeValue(loActivityObjDef, Constants.AT_LINK, "");

                var lmAssignedFAD = getAssignedModel(laObjOcc[i], 14);

                if (lmAssignedFAD != null) {
                    var loOccurrenceOnFAD = laObjOcc[i].ObjDef().OccList([lmAssignedFAD]);
                    if (loOccurrenceOnFAD != null) {

                        // find role 'Responsible'
                        laResultsResponsible = getRelatedOccurrence(loOccurrenceOnFAD[0], Constants.CT_EXEC_2, Constants.OT_PERS_TYPE, laResultsResponsible); // is responsible for	218	CT_EXEC_2 Person type (78)	OT_PERS_TYPE
                        lsResponsibleName = ObjsListNamesToString(laResultsResponsible);

                        // find role 'Accountable'
                        laResultsAccountable = getRelatedOccurrence(loOccurrenceOnFAD[0], Constants.CT_IS_TECH_RESP_3, Constants.OT_PERS_TYPE, laResultsAccountable); // is accountable for	220	CT_IS_TECH_RESP_3 Person type (78)	OT_PERS_TYPE
                        lsAccountableName = ObjsListNamesToString(laResultsAccountable);

                        // find role 'Supported'
                        laResultsSupported = getRelatedOccurrence(loOccurrenceOnFAD[0], Constants.CT_DECD_ON, Constants.OT_PERS_TYPE, laResultsSupported); // is supporting	323	CT_DECD_ON Person type (78)	OT_PERS_TYPE
                        lsSupportedName = ObjsListNamesToString(laResultsSupported);

                        // find role 'Consulted'
                        laResultsConsulted = getRelatedOccurrence(loOccurrenceOnFAD[0], Constants.CT_HAS_CONSLT_ROLE_IN_2, Constants.OT_PERS_TYPE, laResultsConsulted); // is consulted for	358	CT_HAS_CONSLT_ROLE_IN_2 Person type (78)	OT_PERS_TYPE
                        lsConsultedName = ObjsListNamesToString(laResultsConsulted);

                        // find role 'Informed'
                        laResultsInformed = getRelatedOccurrence(loOccurrenceOnFAD[0], Constants.CT_MUST_BE_INFO_ABT_2, Constants.OT_PERS_TYPE, laResultsInformed); // is informed of	326	CT_MUST_BE_INFO_ABT_2 Person type (78)	OT_PERS_TYPE
                        lsInformedName = ObjsListNamesToString(laResultsInformed);

                        // find Application System Types
                        var laResultsApplSysType = getRelatedOccurrence(loOccurrenceOnFAD[0], Constants.CT_CAN_SUPP_1, Constants.OT_APPL_SYS_TYPE, null); // supports	221	CT_CAN_SUPP_1 Application System Type (6) API OT_APPL_SYS_TYPE
                        lsApplicationSystemTypeName = ObjsListNamesToString(laResultsApplSysType);

                        // find 'Screen' Cxn: 73 (is represented by) CT_IS_REPR_BY
                        var laResultsScreens = getRelatedOccurrence(loOccurrenceOnFAD[0], Constants.CT_IS_REPR_BY, Constants.OT_SCRN, null); // is informed of	326	CT_MUST_BE_INFO_ABT_2 Person type (78)	OT_PERS_TYPE
                        for each(resultScreen in laResultsScreens) {
                            var lsScreenName = resultScreen.Name(giLoc);
                            lsScreenFullName = getItemAttributeValue(resultScreen, Constants.AT_NAME_FULL, "");
                            printRow(psModelName, lsActivityName, lsActivityDescription, "", lsOutput, lsInputOutputDescription, lsApplicationSystemTypeName, lsResponsibleName, lsAccountableName, lsSupportedName, lsConsultedName, lsInformedName, lsTitle1, lsLink1, lsTitle2, lsLink2, lsTitle3, lsLink3, lsTitle4, lsLink4, lsScreenName, lsScreenFullName, pExcelWorkBook, rowColorBr);
                            funcBr++;
                        }

                        // find outputs
                        var laResultsOutputs = getRelatedOccurrence(loOccurrenceOnFAD[0], Constants.CT_HAS_OUT, Constants.OT_TECH_TRM, null); // CT_HAS_OUT Technical Term (58) API OT_TECH_TRM
                        for each(resultOutput in laResultsOutputs) {
                            lsOutput = resultOutput.Name(giLoc);
                            lsInputOutputDescription = getItemAttributeValue(resultOutput, 9, "");
                            printRow(psModelName, lsActivityName, lsActivityDescription, "", lsOutput, lsInputOutputDescription, lsApplicationSystemTypeName, lsResponsibleName, lsAccountableName, lsSupportedName, lsConsultedName, lsInformedName, lsTitle1, lsLink1, lsTitle2, lsLink2, lsTitle3, lsLink3, lsTitle4, lsLink4, lsSAPNames, lsSAPFullNames, pExcelWorkBook, rowColorBr);
                            funcBr++;
                            // save print status Input/output
                            lsInputOutputFound = true;
                        }

                        // find inputs
                        var laResultsOutputs = getRelatedOccurrence(loOccurrenceOnFAD[0], Constants.CT_IS_INP_FOR, Constants.OT_TECH_TRM, null); // CT_IS_INP_FOR Technical Term (58) API OT_TECH_TRM
                        for each(resultOutput in laResultsOutputs) {
                            lsInput = resultOutput.Name(giLoc);
                            lsInputOutputDescription = getItemAttributeValue(resultOutput, 9, "");
                            printRow(psModelName, lsActivityName, lsActivityDescription, lsInput, "", lsInputOutputDescription, lsApplicationSystemTypeName, lsResponsibleName, lsAccountableName, lsSupportedName, lsConsultedName, lsInformedName, lsTitle1, lsLink1, lsTitle2, lsLink2, lsTitle3, lsLink3, lsTitle4, lsLink4, lsSAPNames, lsSAPFullNames, pExcelWorkBook, rowColorBr);
                            funcBr++;
                            // save print status Input/output
                            lsInputOutputFound = true;
                        }
                        // when NO Input/output has been found and printed, print one row to make sure activities without Input/Outputs are included
                        if (lsInputOutputFound == false) {
                            printRow(psModelName, lsActivityName, lsActivityDescription, lsInput, lsOutput, lsInputOutputDescription, lsApplicationSystemTypeName, lsResponsibleName, lsAccountableName, lsSupportedName, lsConsultedName, lsInformedName, lsTitle1, lsLink1, lsTitle2, lsLink2, lsTitle3, lsLink3, lsTitle4, lsLink4, lsSAPNames, lsSAPFullNames, pExcelWorkBook, rowColorBr);
                            funcBr++;
                        }
                    } // if there is no FAD, print data with already collected data
                    else {
                        // Print data
                        printRow(psModelName, lsActivityName, lsActivityDescription, lsInput, lsOutput, lsInputOutputDescription, lsApplicationSystemTypeName, lsResponsibleName, lsAccountableName, lsSupportedName, lsConsultedName, lsInformedName, lsTitle1, lsLink1, lsTitle2, lsLink2, lsTitle3, lsLink3, lsTitle4, lsLink4, lsSAPNames, lsSAPFullNames, pExcelWorkBook, rowColorBr);
                        funcBr++;
                    }
                } else {
                    // Print data
                    printRow(psModelName, lsActivityName, lsActivityDescription, lsInput, lsOutput, lsInputOutputDescription, lsApplicationSystemTypeName, lsResponsibleName, lsAccountableName, lsSupportedName, lsConsultedName, lsInformedName, lsTitle1, lsLink1, lsTitle2, lsLink2, lsTitle3, lsLink3, lsTitle4, lsLink4, lsSAPNames, lsSAPFullNames, pExcelWorkBook, rowColorBr);
                    funcBr++;
                }
                //added 04/07/2019
                rowColorBr++;
                //merge cells after print
                mergeCells(funcBr, pExcelWorkBook)
            }
        }
    } catch (ex) {
        Dialogs.MsgBox("Selection invalid, more then one model selected!" + ex);
    }
}
// remove line break from String parameter psText
function filterCarriageReturn(psText) {
    var lsResult = "";
    if (psText != null && psText.equalsIgnoreCase("") == false) {
        lsResult = psText.replaceAll("\\r|\\n", "");
    }
    return lsResult;
}
/*  
 *   Desc: finds assigned model that is of model type number : piModelTypeNum
 */
function getAssignedModel(pObjOcc, piModelTypeNum) {
    var result = null;
    var lbResultFound = false;
    if (pObjOcc.ObjDef().AssignedModels().length > 0) {
        var laAssignedModels = pObjOcc.ObjDef().AssignedModels();
        // while there are assigned models and result has not been found.
        for (var i = 0; i < laAssignedModels.length && lbResultFound == false; i++) {
            // check model type number
            if (laAssignedModels[i].TypeNum() == piModelTypeNum) {
                result = laAssignedModels[i];
                lbResultFound = true;
            }
        }
    }
    return result;
}
/*  
 *   Desc: Collects related object occurrences and merges resultset with paMergeResult
 *   before returning result. paMergeResult is used to store specific related objects before outputing
 *   result to Excel.
 *
 *   pObjOcc: occurrence object that is used for collecting related occurrences
 *   piCnxTypeNum : the connection type number
 *
 */
function getRelatedOccurrence(pObjOcc, piCnxTypeNum, piTargetTypeNum, paMergeResult) {
    var laFoundRelatedObjects = new Array();
    // if parameter array contains objects, use it for collection results
    if (paMergeResult != null) {
        laFoundRelatedObjects = paMergeResult;
    }
    try {
        var lCnxOccList = null;
        lCnxOccList = pObjOcc.Cxns();
        for each(connectionOcc in lCnxOccList) {
            if (connectionOcc.CxnDef().TypeNum() == piCnxTypeNum) {
                var targetOcc = connectionOcc.SourceObjOcc();
                // because direction is not being evaluated, see if pObjOcc is not SourceObjOcc
                if (targetOcc.ObjDef().Name(giLoc).equalsIgnoreCase(pObjOcc.ObjDef().Name(giLoc)) == false) {
                    // check target Obj Occ type number
                    if (targetOcc.ObjDef().TypeNum() == piTargetTypeNum) {

                        laFoundRelatedObjects[laFoundRelatedObjects.length] = targetOcc.ObjDef();
                    }
                } else {
                    targetOcc = connectionOcc.TargetObjOcc();
                    if (targetOcc.ObjDef().Name(giLoc).equalsIgnoreCase(pObjOcc.ObjDef().Name(giLoc)) == false) {
                        // check target Obj Occ type number
                        if (targetOcc.ObjDef().TypeNum() == piTargetTypeNum) {
                            laFoundRelatedObjects[laFoundRelatedObjects.length] = targetOcc.ObjDef();
                        }
                    }
                }
            }
        }
    } catch (ex) {
        Dialogs.MsgBox("Error occured in getRelatedOccurrence." + ex);
    }
    return laFoundRelatedObjects;
}

function getRelatedObjOccurrences(pModel, pObjOcc, piCnxTypeNum, piTargetTypeNum) {
    var laFoundRelatedObjects = new Array();
    try {
        var lCnxDefList = null;
        lCnxDefList = pObjDef.CxnListFilter(piTargetTypeNum);
        for each(connectionOcc in lCnxDefList) {
            if (connectionOcc.TypeNum() == piCnxTypeNum) {
                var loTargetObjDef = connectionOcc.SourceObjDef();
                // when SourceObject is not equal to name of pObjDef object name
                if (loTargetObjDef != null && pObjDef.Name(giLoc).toString().equalsIgnoreCase(loTargetObjDef.Name(giLoc).toString()) == false) {
                    laFoundRelatedObjects[laFoundRelatedObjects.length] = loTargetObjDef;
                } else {
                    loTargetObjDef = connectionOcc.TargetObjDef();
                    if (loTargetObjDef != null && pObjDef.Name(giLoc).toString().equalsIgnoreCase(loTargetObjDef.Name(giLoc).toString()) == false) {
                        laFoundRelatedObjects[laFoundRelatedObjects.length] = loTargetObjDef;
                    }
                }
            }
        }
    } catch (ex) {
        Dialogs.MsgBox("Error occured searching Business Roles." + ex);
    }
    return laFoundRelatedObjects;
}

function getRelatedObjDefs(pObjDef, piCnxTypeNum, piTargetTypeNum) {
    var laFoundRelatedObjects = new Array();
    try {
        var lCnxDefList = null;
        lCnxDefList = pObjDef.CxnListFilter(piTargetTypeNum);
        for each(connectionOcc in lCnxDefList) {
            if (connectionOcc.TypeNum() == piCnxTypeNum) {
                var loTargetObjDef = connectionOcc.SourceObjDef();
                // when SourceObject is not equal to name of pObjDef object name
                if (loTargetObjDef != null && pObjDef.Name(giLoc).toString().equalsIgnoreCase(loTargetObjDef.Name(giLoc).toString()) == false) {
                    laFoundRelatedObjects[laFoundRelatedObjects.length] = loTargetObjDef;
                } else {
                    loTargetObjDef = connectionOcc.TargetObjDef();
                    if (loTargetObjDef != null && pObjDef.Name(giLoc).toString().equalsIgnoreCase(loTargetObjDef.Name(giLoc).toString()) == false) {
                        laFoundRelatedObjects[laFoundRelatedObjects.length] = loTargetObjDef;
                    }
                }
            }
        }
    } catch (ex) {
        Dialogs.MsgBox("Error occured searching Business Roles." + ex);
    }
    return laFoundRelatedObjects;
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

function ObjListToString(paValuesList) {
    var psResult = "";
    var liCounter = 0;
    var nrOfElements = paValuesList.length;
    for each(lsObjStrValue in paValuesList) {
        if (liCounter + 1 == nrOfElements) {
            psResult += gsBulletSign + " " + lsObjStrValue;
        } else {
            psResult += gsBulletSign + " " + lsObjStrValue + "\n";
        }
    }
    return psResult;
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

function printRow(psModelName, psActivityName, psActivityDescription, psInput, psOutput, psInputOutputDescription, psApplicationSystemTypeName, psResponsibleName, psAccountableName, psSupportedName, psConsultedName, psInformedName, psTitle1, psLink1, psTitle2, psLink2, psTitle3, psLink3, psTitle4, psLink4, psSAPNames, psSAPFullNames, pExcelWorkBook, rowColorBr) {
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

        // Ouput Name
        var lCellBPO = lModelRow.createCell(4);
        lCellBPO.setCellValue(psOutput);
        lCellBPO.setCellStyle(crDataRowStyle);

        // Input/Output Description
        var lCellDesc = lModelRow.createCell(5);
        lCellDesc.setCellValue(psInputOutputDescription);
        lCellDesc.setCellStyle(crDataRowStyle);

        // Application system type 
        var lCellDesc = lModelRow.createCell(6);
        lCellDesc.setCellValue(psApplicationSystemTypeName);
        lCellDesc.setCellStyle(crDataRowStyle);

        // Role responsible
        var lCellDesc = lModelRow.createCell(7);
        lCellDesc.setCellValue(psResponsibleName);
        lCellDesc.setCellStyle(crDataRowStyle);

        // Role accountable
        var lCellDesc = lModelRow.createCell(8);
        lCellDesc.setCellValue(psAccountableName);
        lCellDesc.setCellStyle(crDataRowStyle);

        // Role supported
        var lCellDesc = lModelRow.createCell(9);
        lCellDesc.setCellValue(psSupportedName);
        lCellDesc.setCellStyle(crDataRowStyle);

        // Role consulted
        var lCellDesc = lModelRow.createCell(10);
        lCellDesc.setCellValue(psConsultedName);
        lCellDesc.setCellStyle(crDataRowStyle);

        // Role informed
        var lCellDesc = lModelRow.createCell(11);
        lCellDesc.setCellValue(psInformedName);
        lCellDesc.setCellStyle(crDataRowStyle);

        // Title 1
        var lCellDesc = lModelRow.createCell(12);
        lCellDesc.setCellValue(psTitle1);
        lCellDesc.setCellStyle(crDataRowStyle);

        // Link 1
        var lCellDesc = lModelRow.createCell(13);
        lCellDesc.setCellValue(psLink1);
        lCellDesc.setCellStyle(crDataRowStyle);

        // Title 2
        var lCellDesc = lModelRow.createCell(14);
        lCellDesc.setCellValue(psTitle2);
        lCellDesc.setCellStyle(crDataRowStyle);

        // Link 2
        var lCellDesc = lModelRow.createCell(15);
        lCellDesc.setCellValue(psLink2);
        lCellDesc.setCellStyle(crDataRowStyle);

        // Title 3
        var lCellDesc = lModelRow.createCell(16);
        lCellDesc.setCellValue(psTitle3);
        lCellDesc.setCellStyle(crDataRowStyle);

        // Link 3
        var lCellDesc = lModelRow.createCell(17);
        lCellDesc.setCellValue(psLink3);
        lCellDesc.setCellStyle(crDataRowStyle);

        // Title 4
        var lCellDesc = lModelRow.createCell(18);
        lCellDesc.setCellValue(psTitle4);
        lCellDesc.setCellStyle(crDataRowStyle);

        // Link 4
        var lCellDesc = lModelRow.createCell(19);
        lCellDesc.setCellValue(psLink4);
        lCellDesc.setCellStyle(crDataRowStyle);

        //changed 08/07/2019
        // SAP Full Names
        var lCellDesc = lModelRow.createCell(20);
        lCellDesc.setCellValue(psSAPFullNames);
        lCellDesc.setCellStyle(crDataRowStyle);
        
        // SAP Names
        var lCellDesc = lModelRow.createCell(21);
        lCellDesc.setCellValue(psSAPNames);
        lCellDesc.setCellStyle(crDataRowStyle);

        // up counter
        giRowCounter++;
    } catch (ex) {
        Dialogs.MsgBox("Error : " + ex);
    }
}

function printSheetRowModel(lsModelName, pExcelWorkBook) {

    try {

        var lMainSheet = pExcelWorkBook.getSheetAt(0);

        // second row data
        var lModelRow = lMainSheet.createRow(giRowCounter);
        lDataRowStyle.setWrapText(true);

        var lCell0 = lModelRow.createCell(0);
        lCell0.setCellValue(lsModelLevel1Name);
        lDataRowStyle.setWrapText(true);
        lCell0.setCellStyle(lDataRowStyle);

        var lCell1 = lModelRow.createCell(1);
        lDataRowStyle.setWrapText(true);
        lCell1.setCellValue(lsModelLevel2Name);
        lCell1.setCellStyle(lDataRowStyle);

        var lCell2 = lModelRow.createCell(2);
        lDataRowStyle.setWrapText(true);
        lCell2.setCellValue(lsModelLevel3Name);
        lCell2.setCellStyle(lDataRowStyle);

        var lCell3 = lModelRow.createCell(3);
        lCell3.setCellValue(lsModelLevel4Name);
        lDataRowStyle.setWrapText(true);
        lCell3.setCellStyle(lDataRowStyle);

        // BPO Delegate
        var lCellBPO = lModelRow.createCell(5);
        lCellBPO.setCellValue(psPBO);
        lDataRowStyle.setWrapText(true);
        lCellBPO.setCellStyle(lDataRowStyle);

        // Description
        var lCellDesc = lModelRow.createCell(6);
        lCellDesc.setCellValue(psDescription);
        lDataRowStyle.setWrapText(true);
        lCellDesc.setCellStyle(lDataRowStyle);

        // Business process role
        var lCellBusinessProcessRole = lModelRow.createCell(7);
        lCellBusinessProcessRole.setCellValue(psBusinessProcessRole);
        lDataRowStyle.setWrapText(true);
        lCellBusinessProcessRole.setCellStyle(lDataRowStyle);

        // Date created
        var lCellDateCreated = lModelRow.createCell(8);
        lCellDateCreated.setCellValue(psDateCreated);
        lDataRowStyle.setWrapText(true);
        lCellDateCreated.setCellStyle(lDataRowStyle);

        // Date changed
        var lCellDateChanged = lModelRow.createCell(9);
        lCellDateChanged.setCellValue(lsDateChanged);
        lDataRowStyle.setWrapText(true);
        lCellDateChanged.setCellStyle(lDataRowStyle);

        // Document name
        var lCellDocumentName = lModelRow.createCell(10);
        lCellDocumentName.setCellValue("");
        lDataRowStyle.setWrapText(true);
        lCellDocumentName.setCellStyle(lDataRowStyle);

        // Document links
        var lCellDocumentLinks = lModelRow.createCell(11);
        lCellDocumentLinks.setCellValue(psLinksCellValue);
        lDataRowStyle.setWrapText(true);
        lCellDocumentLinks.setCellStyle(lDataRowStyle);

        // up counter
        giRowCounter++;
    } catch (ex) {
        Dialogs.MsgBox("Error : " + ex);
    }
}

function isValidActivity(pObjOccurence) {
    var lbActivityValid = false;
    var liSymbolTypeNum = pObjOccurence.getSymbol();
    if (Constants.ST_FUNC == liSymbolTypeNum) {
        lbActivityValid = true;
    } else if (Constants.ST_PRCS_IF == liSymbolTypeNum) {
        lbActivityValid = false;
    }
    return lbActivityValid;
}

function formatLink(psTitle, psLink) {
    var lsResult = "";
    if (psTitle.equalsIgnoreCase("") == false) {
        // remove line breaks
        psTitle = filterCarriageReturn(psTitle);
        if (psLink.equalsIgnoreCase("") == false) {
            lsResult = gsBulletSign + " " + psTitle + " (" + psLink + ")";
        } else {
            lsResult = gsBulletSign + " " + psTitle;

        }
    } else if (psLink.equalsIgnoreCase("") == false) {
        lsResult = gsBulletSign + " " + psLink;
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
        regionSheet.addRegion(6, 6, fromRow, giRowCounter - 1)
        regionSheet.addRegion(7, 7, fromRow, giRowCounter - 1)
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
    }
}

main()