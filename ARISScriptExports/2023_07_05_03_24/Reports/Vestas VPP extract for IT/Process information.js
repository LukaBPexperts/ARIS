var giLoc = Context.getSelectedLanguage();
var gExcelObject = null;
var giStandardColumnWidth=15000;
var giRowCounter = 0;
var lDataRowStyle = null;
var gCurrentFilter;

var gaModelsL1Process;
var gaModelsL2Process;
var gaModelsL3Process;
var gaModelsL4Process;

var ghmModelsL1Process = new java.util.HashMap();
var ghmModelsL2Process = new java.util.HashMap();
var ghmModelsL3Process = new java.util.HashMap();
var ghmModelsL4Process = new java.util.HashMap();

var ghsModelsPassed = new java.util.HashSet();

const giProcessFlowChartStandard = Constants.MT_EEPC; // Process Flow Chart 
const giProcessFlowChartStandardSwimlanes = Constants.MT_EEPC_COLUMN; // Process Flow Chart - Swimlanes (EPC (column display))
const giProcessLevelAttribute =  ArisData.ActiveFilter().UserDefinedAttributeTypeNum("32a18230-aaa7-11df-4732-00237d347a8a");
const giStatusReleaseAttribute =  ArisData.ActiveFilter().UserDefinedAttributeTypeNum("8fda52b0-6188-11de-4732-00237d347a8a"); // status
const gsApprovedReleasedStatus = "Released";

/*
*   JavaScript class to store Model path to L0
*/
var tModelPath = function tModelPath(psGUID, piLevel){
    this.modelName = "";
    this.GUID = psGUID;
    this.L1_model_name = "";
    this.L2_model_name = "";
    this.L3_model_name = "";
    this.L4_model_name = "";
    this.modelLevel = piLevel;
    this.partOfSipoc = false;
}


// Constants
const gsBPOAttrGUID = "618d4120-ea3f-11de-4732-00237d347a8a";

function main(){
    gExcelObject = Context.createExcelWorkbook(Context.getSelectedFile()); 
    if(ArisData.getSelectedDatabases().length==1){
        // save active metod filter
        gCurrentFilter = ArisData.ActiveFilter();
        printRowHeaders(gExcelObject);
        var modelsFound = getProcessModels();
        if (modelsFound==true){
            for each(model in gaModelsL1Process){
                printData(model, gExcelObject, 1);
            }
            for each(model in gaModelsL2Process){
                printData(model, gExcelObject, 2);
            }
            for each(model in gaModelsL3Process){
                printData(model, gExcelObject, 3);
            }
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
}

function getProcessModels(){
    try{
        gaModelsL1Process = ArisData.getActiveDatabase().Find(Constants.SEARCH_MODEL, [giProcessFlowChartStandard, giProcessFlowChartStandardSwimlanes], giProcessLevelAttribute, giLoc, 1, Constants.SEARCH_CMP_EQUAL);
        for each(model in gaModelsL1Process){
            ghmModelsL1Process.put(model.GUID(),model); 
        }
        gaModelsL2Process = ArisData.getActiveDatabase().Find(Constants.SEARCH_MODEL, [giProcessFlowChartStandard], giProcessLevelAttribute, giLoc, 2, Constants.SEARCH_CMP_EQUAL);
        for each(model in gaModelsL2Process){
            ghmModelsL2Process.put(model.GUID(),model);         
            getObjProcessHierarchyPath(model, 2);
        }
        gaModelsL3Process = ArisData.getActiveDatabase().Find(Constants.SEARCH_MODEL, [giProcessFlowChartStandard], giProcessLevelAttribute, giLoc, 3, Constants.SEARCH_CMP_EQUAL);
        for each(model in gaModelsL3Process){     
            getObjProcessHierarchyPath(model, 3);
        }
        gaModelsL4Process = ArisData.getActiveDatabase().Find(Constants.SEARCH_MODEL, [giProcessFlowChartStandardSwimlanes]);
        for each(model in gaModelsL4Process){
            ghmModelsL4Process.put(model.GUID(),model);     
            getObjProcessHierarchyPath(model, 4);
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
    // set column width
    // L1
    lMainSheet.setColumnWidth (0, giStandardColumnWidth);
    // L2
    lMainSheet.setColumnWidth (1, giStandardColumnWidth);
    // L3
    lMainSheet.setColumnWidth (2, giStandardColumnWidth);
    // L4
    lMainSheet.setColumnWidth (3, giStandardColumnWidth);
    // Activity Name
    lMainSheet.setColumnWidth (4, giStandardColumnWidth);
    // BPO Delegate
    lMainSheet.setColumnWidth (5, giStandardColumnWidth);
    // Description
    lMainSheet.setColumnWidth (6, giStandardColumnWidth);
    // Business process role
    lMainSheet.setColumnWidth (7, giStandardColumnWidth);
    // Date created
    lMainSheet.setColumnWidth (8, (giStandardColumnWidth/2));
    // Date changed
    lMainSheet.setColumnWidth (9, (giStandardColumnWidth/2));
    // Document name
    lMainSheet.setColumnWidth (10, giStandardColumnWidth);
    // Document links
    var liDocLinkColumnWidth = giStandardColumnWidth * 4;
    lMainSheet.setColumnWidth (11, liDocLinkColumnWidth);
    
    /* Freeze top row alone */
    lMainSheet.createFreezePane(0,1);
    
    // set style
    var lHeaderFont = pExcelWorkBook.createFont();
    lHeaderFont.setBoldweight(Constants.XL_FONT_WEIGHT_BOLD);
    
    
    var lCellL1 = lHeaderRow.createCell(0);
    lCellL1.setCellValue("L1");
    var lCellStyle = lCellL1.getCellStyle();
    lCellStyle.setFont(lHeaderFont);
    var lCellL2 = lHeaderRow.createCell(1);
    lCellL2.setCellValue("L2");
    var lCellL3 = lHeaderRow.createCell(2);
    lCellL3.setCellValue("L3");
    var lCellL4 = lHeaderRow.createCell(3);
    lCellL4.setCellValue("L4");
    var lCell5 = lHeaderRow.createCell(4);
    lCell5.setCellValue("Activity Name");
    var lCell6 = lHeaderRow.createCell(5);
    lCell6.setCellValue("BPO Delegate");
    var lCell7 = lHeaderRow.createCell(6);
    lCell7.setCellValue("Description");
    var lCell8 = lHeaderRow.createCell(7);
    lCell8.setCellValue("Business process role");
    var lCell9 = lHeaderRow.createCell(8);
    lCell9.setCellValue("Date created");
    var lCell10 = lHeaderRow.createCell(9);
    lCell10.setCellValue("Date changed");
    var lCell11 = lHeaderRow.createCell(10);
    lCell11.setCellValue("Document name");
    var lCell12 = lHeaderRow.createCell(11);
    lCell12.setCellValue("Document links");
    
    // set new style
    var lDataFont = pExcelWorkBook.createFont();
    lDataFont.setBoldweight(Constants.XL_FONT_WEIGHT_NORMAL);
      
    var liBorderColor = Constants.C_GRAY;
    lDataRowStyle = pExcelWorkBook.createCellStyle(lDataFont, 1, 1, 1, 1, liBorderColor, liBorderColor, liBorderColor, liBorderColor, Constants.ALIGN_LEFT, Constants.VERTICAL_TOP, Constants.C_TRANSPARENT,Constants.C_TRANSPARENT,Constants.SOLID_FOREGROUND);
    lDataRowStyle.setWrapText(true);
    // set style
    var lCell13 = lHeaderRow.createCell(12);
    lCell13.setCellStyle(lDataRowStyle);
    lCell13.setCellValue(" ");
    
    giRowCounter++
}
function isSIPOC(pObjOccurence){
    var lbSIPOCValid = false;
    var liSymbolTypeNum = pObjOccurence.getSymbol();
    
    const lsSIPOC = "ff856510-b087-11df-4732-00237d347a8a";
    const lsSIPOC2 = "693974af-b0e0-11df-4732-00237d347a8a";
    
    if(getCustomSymbolTypeNum(lsSIPOC) == liSymbolTypeNum){
        lbSIPOCValid = true;
    }else if(getCustomSymbolTypeNum(lsSIPOC2) == liSymbolTypeNum){
        lbSIPOCValid = true;
    }    
    return lbSIPOCValid;
}
function isValidActivity(pObjOccurence){
    var lbActivityValid = false;
    var liSymbolTypeNum = pObjOccurence.getSymbol();
    
    const lsE2EInterfaceSap = "9bc05050-3ad0-11e7-004f-eabd1a8bb6a2";
    const lsE2EInterfaceStandard = "87d697d0-1878-11e8-67df-0800274d6e84";
    const lsSIPOC = "ff856510-b087-11df-4732-00237d347a8a";
    const lsSIPOC2 = "693974af-b0e0-11df-4732-00237d347a8a";
    const lsAutomatedActivity = "2a849d90-d12a-11df-4732-00237d347a8a";
    
    if(Constants.ST_FUNC == liSymbolTypeNum){
        lbActivityValid = true;
    }else if(Constants.ST_SAP_NETWEAV_BUSINESS_PROC == liSymbolTypeNum){
        lbActivityValid = true;
    }else if(getCustomSymbolTypeNum(lsE2EInterfaceSap) == liSymbolTypeNum){
        lbActivityValid = true;
    }else if(getCustomSymbolTypeNum(lsE2EInterfaceStandard) == liSymbolTypeNum){
        lbActivityValid = true;
    }else if(getCustomSymbolTypeNum(lsSIPOC) == liSymbolTypeNum){
        lbActivityValid = true;
    }else if(getCustomSymbolTypeNum(lsSIPOC2) == liSymbolTypeNum){
        lbActivityValid = true;
    }else if(getCustomSymbolTypeNum(lsAutomatedActivity) == liSymbolTypeNum){
        lbActivityValid = true;
    }    
    return lbActivityValid;
}

function isValidSuperiorObject(pObjOccurence){
    var lbIsValidSuperiorObject = false;
    var liSymbolTypeNum = pObjOccurence.getSymbol();
    
    const lsSIPOC = "ff856510-b087-11df-4732-00237d347a8a";
    const lsSIPOC2 = "693974af-b0e0-11df-4732-00237d347a8a";
    const lsAutomatedActivity = "2a849d90-d12a-11df-4732-00237d347a8a";
    
    if(Constants.ST_FUNC == liSymbolTypeNum){
        lbIsValidSuperiorObject = true;
    }else if(Constants.ST_SAP_NETWEAV_BUSINESS_PROC == liSymbolTypeNum){
        lbIsValidSuperiorObject = true;
    }else if(getCustomSymbolTypeNum(lsSIPOC) == liSymbolTypeNum){
        lbIsValidSuperiorObject = true;
    }else if(getCustomSymbolTypeNum(lsSIPOC2) == liSymbolTypeNum){
        lbIsValidSuperiorObject = true;
    }else if(getCustomSymbolTypeNum(lsAutomatedActivity) == liSymbolTypeNum){
        lbIsValidSuperiorObject = true;
    }    
    return lbIsValidSuperiorObject;
}
function isValidActivity(pObjOccurence){
    var lbActivityValid = false;
    var liSymbolTypeNum = pObjOccurence.getSymbol();
    
    const lsE2EInterfaceSap = "9bc05050-3ad0-11e7-004f-eabd1a8bb6a2";
    const lsE2EInterfaceStandard = "87d697d0-1878-11e8-67df-0800274d6e84";
    const lsSIPOC = "ff856510-b087-11df-4732-00237d347a8a";
    const lsSIPOC2 = "693974af-b0e0-11df-4732-00237d347a8a";
    const lsAutomatedActivity = "2a849d90-d12a-11df-4732-00237d347a8a";
    
    if(Constants.ST_FUNC == liSymbolTypeNum){
        lbActivityValid = true;
    }else if(Constants.ST_SAP_NETWEAV_BUSINESS_PROC == liSymbolTypeNum){
        lbActivityValid = true;
    }else if(getCustomSymbolTypeNum(lsE2EInterfaceSap) == liSymbolTypeNum){
        lbActivityValid = true;
    }else if(getCustomSymbolTypeNum(lsE2EInterfaceStandard) == liSymbolTypeNum){
        lbActivityValid = true;
    }else if(getCustomSymbolTypeNum(lsSIPOC) == liSymbolTypeNum){
        lbActivityValid = true;
    }else if(getCustomSymbolTypeNum(lsSIPOC2) == liSymbolTypeNum){
        lbActivityValid = true;
    }else if(getCustomSymbolTypeNum(lsAutomatedActivity) == liSymbolTypeNum){
        lbActivityValid = true;
    }    
    return lbActivityValid;
}
//  checks if model has 'Status':'Releaed
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
/* 
*   checks is model has expeced 'Level Graphic' attribute maintained
*   and if attribute value is as expected
*/
function modelHasExpectedModelGraphic(pModel, piExpectedLevelGraphic){
    var lbResultIsExpectedLevelGraphic=false;
    try{
        var lModelGraphicAttribute = pModel.Attribute(giStatusReleaseAttribute,giLoc);
        if (lModelGraphicAttribute.IsMaintained()==true 
        && lModelGraphicAttribute.getValue().toString().equalsIgnoreCase(java.lang.Integer.toString(piExpectedLevelGraphic))==true){
            lbResultIsExpectedLevelGraphic=true;
        }
    }catch(ex){        
        Dialogs.MsgBox("Error while checking if model attribute 'Level Graphic' contains value '" + piExpectedLevelGraphic + "'." + ex);
    }
    return lbResultIsExpectedLevelGraphic;
}
function printData(pModel, pExcelWorkBook, piModelLevel){
    try{
        // don't process models that are don't have the attribute status 'Released'
        if (modelIsReleased(pModel)==true){
            if (ghsModelsPassed.contains("" + pModel.GUID())==false){
                
                var lsModelPrintData = true;
                ghsModelsPassed.add("" + pModel.GUID());                   
                
                var lsModelLevel1Name = "";
                var lsModelLevel2Name = "";
                var lsModelLevel3Name = "";
                var lsModelLevel4Name = "";
              
                if (piModelLevel==0){
                    lsModelLevel1Name = pModel.Name(giLoc);
                }
                
                if (piModelLevel==1){
                    var tTModelPath = ghmModelsL1Process.get(pModel.GUID());
                    if (tTModelPath!=null){
                        lsModelLevel1Name = tTModelPath.L1_model_name;
                    }
                    lsModelLevel2Name = pModel.Name(giLoc);
                }
                if (piModelLevel==2){
                    var tTModelPath = ghmModelsL2Process.get(pModel.GUID());
                    if (tTModelPath!=null){
                        lsModelLevel1Name = tTModelPath.L1_model_name;
                        lsModelLevel2Name = tTModelPath.L2_model_name;
                    }
                    lsModelLevel3Name = pModel.Name(giLoc);
                }
                if (piModelLevel==3){
                    var tTModelPath = ghmModelsL3Process.get(pModel.GUID());
                    if (tTModelPath!=null){
                        lsModelLevel1Name = tTModelPath.L1_model_name;
                        lsModelLevel2Name = tTModelPath.L2_model_name;
                        lsModelLevel3Name = tTModelPath.L3_model_name;
                    }else{
                        tTModelPath = ghmModelsL2Process.get(pModel.GUID());
                        if (tTModelPath!=null){
                        lsModelLevel1Name = tTModelPath.L1_model_name;
                        lsModelLevel2Name = tTModelPath.L2_model_name;
                        }
                    }
                    lsModelLevel4Name = pModel.Name(giLoc);
                }
                if (piModelLevel==4){
                    var tTModelPath = ghmModelsL4Process.get(pModel.GUID());
                    if (tTModelPath!=null){
                        lsModelLevel1Name = tTModelPath.L1_model_name;
                        lsModelLevel2Name = tTModelPath.L2_model_name;
                        lsModelLevel3Name = tTModelPath.L3_model_name;
                    }
                    lsModelLevel4Name = pModel.Name(giLoc);
                }
            
                var sName = pModel.Name(giLoc);
                // BPO
                var liBPOAttTypeNum = ArisData.ActiveFilter().UserDefinedAttributeTypeNum(gsBPOAttrGUID);
    
                var sPBO = getItemAttributeValue(pModel, liBPOAttTypeNum, "n/a")
                var lsDescription = getItemAttributeValue(pModel, 9, "");
                var lsDateCreated = getItemAttributeValue(pModel, 1008, "");
                var lsDateChanged = getItemAttributeValue(pModel, 1175, "");
                
                var lsTitle1 = getItemAttributeValue(pModel, Constants.AT_TITL1,"");
                var lsTitle2 = getItemAttributeValue(pModel, Constants.AT_TITL2,"");
                var lsTitle3 = getItemAttributeValue(pModel, Constants.AT_TITL3,"");
                var lsTitle4 = getItemAttributeValue(pModel, Constants.AT_TITL4,"");
                
               var lsLink1 = getItemAttributeValue(pModel, Constants.AT_EXT_1,"");
                var lsLink2 = getItemAttributeValue(pModel, Constants.AT_EXT_2,"");
                var lsLink3 = getItemAttributeValue(pModel, Constants.AT_EXT_3,"");
                var lsLink4 = getItemAttributeValue(pModel, Constants.AT_LINK,"");
    
                var lsModelLinksCellValue = prepareLinkCell(lsTitle1, lsLink1, lsTitle2, lsLink2, lsTitle3, lsLink3, lsTitle4, lsLink4);
    
                var liModelLevel = 0;
                var lsBusinessProcessRole = "";           
                var laObjOcc = pModel.ObjOccList();
                var lsLinksCellValue = "";
                if (piModelLevel!=0)
                for (var i=0 ; i < laObjOcc.length ; i++){
                    // Activity (Check for Process interfaces)
                    if (laObjOcc[i].ObjDef().TypeNum().toString().equalsIgnoreCase("22") 
                        && isValidActivity(laObjOcc[i])==true){
    
                        var lsActivityName = laObjOcc[i].ObjDef().Name(giLoc);
                        var loActivityDef = laObjOcc[i].ObjDef();
                        var lsActivityDescription = getItemAttributeValue(loActivityDef, 9, "");
                        var laBusinessRoles = getRelatesBusinessRoles(laObjOcc[i]);
                                
                        lsDateCreated = getItemAttributeValue(loActivityDef, 1008, "");
                        lsDateChanged = getItemAttributeValue(loActivityDef, 1175, "");
                
                        lsTitle1 = getItemAttributeValue(loActivityDef, Constants.AT_TITL1,"");
                        lsTitle2 = getItemAttributeValue(loActivityDef, Constants.AT_TITL2,"");
                        lsTitle3 = getItemAttributeValue(loActivityDef, Constants.AT_TITL3,"");
                        lsTitle4 = getItemAttributeValue(loActivityDef, Constants.AT_TITL4,"");
                        
                        lsLink1 = getItemAttributeValue(loActivityDef, Constants.AT_EXT_1,"");
                        lsLink2 = getItemAttributeValue(loActivityDef, Constants.AT_EXT_2,"");
                        lsLink3 = getItemAttributeValue(loActivityDef, Constants.AT_EXT_3,"");
                        lsLink4 = getItemAttributeValue(loActivityDef, Constants.AT_LINK,"");
                        
                        lsLinksCellValue = prepareLinkCell(lsTitle1, lsLink1, lsTitle2, lsLink2, lsTitle3, lsLink3, lsTitle4, lsLink4);                    

                        if (piModelLevel==3){
                          
                            if (tTModelPath.L2_model_name.equalsIgnoreCase("")==false){
                                lsModelLevel1Name = tTModelPath.L1_model_name;
                                lsModelLevel2Name = tTModelPath.L2_model_name;
                                lsModelLevel3Name = lsActivityName;
                                lsModelLevel4Name = "";         
                            }else{
                                lsModelLevel1Name = sName;
                                lsModelLevel2Name = lsActivityName;
                                lsModelLevel3Name = lsActivityName;
                                lsModelLevel4Name = "";     
                            }                            
                            printSheetRowModel(lsModelLevel1Name, lsModelLevel2Name, lsModelLevel3Name, lsModelLevel4Name, sPBO, lsDescription, lsBusinessProcessRole, lsDateCreated, lsDateChanged, lsModelLinksCellValue, piModelLevel, pExcelWorkBook);
                        }else if (piModelLevel==2){        
                            lsModelLevel1Name = tTModelPath.L1_model_name;
                            lsModelLevel2Name = lsActivityName;
                            lsModelLevel3Name = "";
                            lsModelLevel4Name = "";
                            printSheetRowModel(lsModelLevel1Name, lsModelLevel2Name, lsModelLevel3Name, lsModelLevel4Name, sPBO, lsDescription, lsBusinessProcessRole, lsDateCreated, lsDateChanged, lsModelLinksCellValue, piModelLevel, pExcelWorkBook);
                        }else if (piModelLevel==4){                                  
                            lsModelLevel1Name = tTModelPath.L1_model_name;
                            lsModelLevel2Name = tTModelPath.L2_model_name;
                            lsModelLevel3Name = tTModelPath.L3_model_name
                            lsModelLevel4Name = tTModelPath.L4_model_name;
                            printSheetRowActivity(lsModelLevel1Name, lsModelLevel2Name, lsModelLevel3Name, lsModelLevel4Name, lsActivityName, sPBO, lsActivityDescription, laBusinessRoles, lsDateCreated, lsDateChanged, lsLinksCellValue, piModelLevel, pExcelWorkBook);
                            // printSheetRowModel(lsModelLevel1Name, lsModelLevel2Name, lsModelLevel3Name, lsModelLevel4Name, sPBO, lsDescription, lsBusinessProcessRole, lsDateCreated, lsDateChanged, lsModelLinksCellValue, piModelLevel, pExcelWorkBook);
                        }
                        
                    }
                    // Document
                    else if(laObjOcc[i].ObjDef().TypeNum().toString().equalsIgnoreCase("27")){
                        var lsDocumentName = laObjOcc[i].ObjDef().Name(giLoc);
                        var loDocumentDef = laObjOcc[i].ObjDef();
                                
                        lsDateCreated = getItemAttributeValue(loDocumentDef, 1008, "");
                        lsDateChanged = getItemAttributeValue(loDocumentDef, 1175, "");
                
                        lsTitle1 = getItemAttributeValue(loDocumentDef, Constants.AT_TITL1,"");
                        lsTitle2 = getItemAttributeValue(loDocumentDef, Constants.AT_TITL2,"");
                        lsTitle3 = getItemAttributeValue(loDocumentDef, Constants.AT_TITL3,"");
                        lsTitle4 = getItemAttributeValue(loDocumentDef, Constants.AT_TITL4,"");
                        
                        lsLink1 = getItemAttributeValue(loDocumentDef, Constants.AT_EXT_1,"");
                        lsLink2 = getItemAttributeValue(loDocumentDef, Constants.AT_EXT_2,"");
                        lsLink3 = getItemAttributeValue(loDocumentDef, Constants.AT_EXT_3,"");
                        lsLink4 = getItemAttributeValue(loDocumentDef, Constants.AT_LINK,"");
                        
                        lsLinksCellValue = prepareLinkCell(lsTitle1, lsLink1, lsTitle2, lsLink2, lsTitle3, lsLink3, lsTitle4, lsLink4);                    
                        
                        printSheetRowDocument(lsModelLevel1Name, lsModelLevel2Name, lsModelLevel3Name, lsModelLevel4Name, lsDocumentName, sPBO, lsDescription, "", lsDateCreated, lsDateChanged, lsLinksCellValue, piModelLevel, pExcelWorkBook);
                    }
                }
                
                if (liModelLevel!=null && lsModelPrintData==true && piModelLevel==4){
                    printSheetRowModel(lsModelLevel1Name, lsModelLevel2Name, lsModelLevel3Name, lsModelLevel4Name, sPBO, lsDescription, lsBusinessProcessRole, lsDateCreated, lsDateChanged, lsModelLinksCellValue, piModelLevel, pExcelWorkBook);
                }
                /*else{
                    Dialogs.MsgBox("Could not determine level for this model! stopping report.");   
                }
                */
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

function getObjProcessHierarchyPath(pModel, piModelLevel){
    var lbResultFound=false;
    try {
        var lModelPath = new tModelPath(pModel.GUID(), piModelLevel);
        if (piModelLevel==2){
            lModelPath.modelName = pModel.Name(giLoc);
            // lModelPath.L2_model_name = pModel.Name(giLoc);              
            var laSupObjList = pModel.getSuperiorObjDefs();
            for each(superiorObjectDef in laSupObjList){
                for each(modelL1 in gaModelsL1Process){
                    var laCheckOccModelList = modelL1.ObjOccList();
                    for each(checkOccModel in laCheckOccModelList){
                        // check is superior object has correct symbol
                        if (isValidSuperiorObject(checkOccModel)==true){
                            for each(superiorObject in laSupObjList){
                                if (checkOccModel.ObjDef().GUID().toString().equalsIgnoreCase(superiorObject.GUID())==true){                            
                                    lModelPath.L1_model_name = checkOccModel.ObjDef().Name(giLoc);
                                    lModelPath.modelName = pModel.Name(giLoc);
                                    lModelPath.GUID = pModel.GUID();                         
                                }
                            }
                        }
                    }                        
                }        
            }        
            ghmModelsL2Process.put(pModel.GUID(),lModelPath);
        }
        else if (piModelLevel==3){
            lModelPath.modelName = pModel.Name(giLoc);
            lModelPath.L3_model_name = pModel.Name(giLoc);              
            var laSupObjList = pModel.getSuperiorObjDefs();
            for each(superiorObjectDef in laSupObjList){
                for each(modelL2 in gaModelsL2Process){
                    var laCheckOccModelList = modelL2.ObjOccList();
                    for each(checkOccModel in laCheckOccModelList){
                        // check is superior object has correct symbol
                        if (isValidSuperiorObject(checkOccModel)==true){
                            for each(superiorObject in laSupObjList){
                                if (checkOccModel.ObjDef().GUID().toString().equalsIgnoreCase(superiorObject.GUID())==true && lbResultFound==false){                            
                                    lModelPath.L2_model_name = modelL2.Name(giLoc);
                                    var superiorPath = ghmModelsL2Process.get(modelL2.GUID());
                                    if (superiorPath!=null){
                                        lModelPath.L1_model_name = superiorPath.L1_model_name;
                                        lModelPath.L2_model_name = checkOccModel.ObjDef().Name(giLoc);
                                        lModelPath.modelName = pModel.Name(giLoc);
                                        lModelPath.GUID = pModel.GUID();                                    
                                        lbResultFound=true;
                                    }
                                }
                            }
                        }
                    }
                    if (lbResultFound==false){
                        var laCheckOccModelList = modelL1.ObjOccList();
                        for each(checkOccModel in laCheckOccModelList){
                            // check is superior object has correct symbol
                            if (isValidSuperiorObject(checkOccModel)==true){
                                for each(superiorObject in laSupObjList){
                                    if (checkOccModel.ObjDef().GUID().toString().equalsIgnoreCase(superiorObject.GUID())==true && lbResultFound==false){                            
                                        //lModelPath.L2_model_name = modelL2.Name(giLoc);
                                        var superiorPath = ghmModelsL1Process.get(modelL1.GUID());
                                        if (superiorPath!=null){
                                            lModelPath.modelName = pModel.Name(giLoc);
                                            lModelPath.L1_model_name = checkOccModel.ObjDef().Name(giLoc);
                                            lModelPath.L2_model_name = pModel.Name(giLoc);
                                            lModelPath.modelName = pModel.Name(giLoc);
                                            lModelPath.GUID = pModel.GUID();
                                            lbResultFound=true;                               
                                        }
                                    }
                                }
                            }
                        }                        
                    }
                }        
            }        
            ghmModelsL3Process.put(pModel.GUID(),lModelPath);
        }
        else if (piModelLevel==4){
            lModelPath.modelName = pModel.Name(giLoc);
            lModelPath.L4_model_name = pModel.Name(giLoc);              
            var laSupObjList = pModel.getSuperiorObjDefs();
            for each(superiorObjectDef in laSupObjList){
                for each(modelL3 in gaModelsL3Process){
                    var laCheckOccModelList = modelL3.ObjOccList();
                    for each(checkOccModel in laCheckOccModelList){
                        // check is superior object has correct symbol
                        if (isValidSuperiorObject(checkOccModel)==true){
                            for each(superiorObject in laSupObjList){
                                if (checkOccModel.ObjDef().GUID().toString().equalsIgnoreCase(superiorObject.GUID())==true && lbResultFound==false){                            
                                    lModelPath.L3_model_name = modelL3.Name(giLoc);
                                    var superiorPath = ghmModelsL3Process.get(modelL3.GUID());
                                    if (superiorPath!=null){
                                        lModelPath.L1_model_name = superiorPath.L1_model_name;
                                        lModelPath.L2_model_name = superiorPath.L2_model_name;
                                        lModelPath.L3_model_name = checkOccModel.ObjDef().Name(giLoc);    
                                        lbResultFound=true;                              
                                    }
                                    if (isSIPOC(checkOccModel)==true){
                                        lModelPath.partOfSipoc=true;   
                                    }
                                }
                            }
                        }
                    }
                }        
            }        
            ghmModelsL4Process.put(pModel.GUID(),lModelPath);
        }
    }
    catch(ex){
         Dialogs.MsgBox("Error occurred!" + ex);
    }
}
function getCustomSymbolTypeNum(psGUID){
    var liSymbolTypeNum = -1;
    if(psGUID.length > 25){
        liSymbolTypeNum = gCurrentFilter.UserDefinedSymbolTypeNum(psGUID);        
    }
    else{
        liSymbolTypeNum = 0;   
    }
    return liSymbolTypeNum;
}
function getRelatesBusinessRoles(pObjOcc){
    var laFoundRoles = new Array();
    try{
        var lCnxOccList = pObjOcc.CxnOccList();
        for each(connectionOcc in lCnxOccList){
            if (connectionOcc.CxnDef().TypeNum()==218){
                var loObjDefSource = connectionOcc.CxnDef().SourceObjDef();
                var loObjDefTarget = connectionOcc.CxnDef().TargetObjDef();
                if (loObjDefTarget.GUID().toString().equalsIgnoreCase(pObjOcc.ObjDef().GUID())==false){
                    laFoundRoles[laFoundRoles.length]=loObjDefTarget.Name(giLoc);
                }else if (loObjDefSource.GUID().toString().equalsIgnoreCase(pObjOcc.ObjDef().GUID())==false){
                    laFoundRoles[laFoundRoles.length]=loObjDefSource.Name(giLoc);
                }
            }  
        }
    }
    catch(ex){
        Dialogs.MsgBox("Error occured searching Business Roles." + ex);
    }
    return laFoundRoles;
}
function prepareLinkCell(psTitle1, psLink1, psTitle2, psLink2, psTitle3, psLink3, psTitle4, psLink4){    
    var lsResultLink = "";

    // link 1
    var lsResultLink1 = formatLink(psTitle1, psLink1);
    // link 2
    var lsResultLink2 = formatLink(psTitle2, psLink2);
    // link 3
    var lsResultLink3 = formatLink(psTitle3, psLink3);
    // link 4
    var lsResultLink4 = formatLink(psTitle4, psLink4);
    
    // when link 1 NOT is empty
    if (lsResultLink1.equalsIgnoreCase("")==false){
        lsResultLink = lsResultLink + lsResultLink1 + "\n";
    }
    if (lsResultLink2.equalsIgnoreCase("")==false){
        lsResultLink = lsResultLink + lsResultLink2 + "\n";
    }
    if (lsResultLink3.equalsIgnoreCase("")==false){
        lsResultLink = lsResultLink + lsResultLink3 + "\n";
    }
    if (lsResultLink4.equalsIgnoreCase("")==false){
        lsResultLink = lsResultLink + lsResultLink4 + "\n";
    }
    
    return lsResultLink;    
}

function formatLink(psTitle, psLink){
    var lsResult="";
    if (psTitle.equalsIgnoreCase("")==false){
        // remove line breaks
        psTitle = filterCarriageReturn(psTitle);
        if(psLink.equalsIgnoreCase("")==false){
            lsResult = "* " + psTitle + " (" + psLink + ")";   
        }else{
           lsResult = "* " + psTitle;  
            
        }
    }else if(psLink.equalsIgnoreCase("")==false){
           lsResult = "* " + psLink;      
    }
    return lsResult;
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

function printSheetRowModel(lsModelLevel1Name, lsModelLevel2Name, lsModelLevel3Name, lsModelLevel4Name, psPBO, psDescription, psBusinessProcessRole, psDateCreated, lsDateChanged, psLinksCellValue, piLevel, pExcelWorkBook){ 
    try{
        
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
    }
    catch(ex){
        Dialogs.MsgBox("Error : " + ex );
    }
}

function printSheetRowActivity(lsModelLevel1Name, lsModelLevel2Name, lsModelLevel3Name, lsModelLevel4Name, psActivityName, psPBO, psActivityDescription, paBusinessProcessRoles, psDateCreated, lsDateChanged, psLinksCellValue, piLevel, pExcelWorkBook){ 
    try{
        
        var lMainSheet = pExcelWorkBook.getSheetAt(0);
        
        // second row data
        var lModelRow = lMainSheet.createRow(giRowCounter);
        var lCell0 = lModelRow.createCell(0);
        lCell0.setCellValue(lsModelLevel1Name);
        lCell0.setCellStyle(lDataRowStyle);
        
        var lCell1 = lModelRow.createCell(1);
        lCell1.setCellValue(lsModelLevel2Name);
        lCell1.setCellStyle(lDataRowStyle);
        
        var lCell2 = lModelRow.createCell(2);
        lCell2.setCellValue(lsModelLevel3Name);
        lCell2.setCellStyle(lDataRowStyle);
        
        var lCell3 = lModelRow.createCell(3);
        lCell3.setCellValue(lsModelLevel4Name);
        lCell3.setCellStyle(lDataRowStyle);        
        
        // Activity Name
        var lCellBPO = lModelRow.createCell(4);
        lCellBPO.setCellValue(psActivityName);
        lCellBPO.setCellStyle(lDataRowStyle);
        
        // BPO Delegate
        var lCellBPO = lModelRow.createCell(5);
        lCellBPO.setCellValue(psPBO);
        lCellBPO.setCellStyle(lDataRowStyle);
        
        // Description
        var lCellDesc = lModelRow.createCell(6);
        lCellDesc.setCellValue(psActivityDescription);
        lCellDesc.setCellStyle(lDataRowStyle);
        
        // Business process role
        var lCellBusinessProcessRole = lModelRow.createCell(7);
        lCellBusinessProcessRole.setCellValue(businessRolesToString(paBusinessProcessRoles));
        lCellBusinessProcessRole.setCellStyle(lDataRowStyle);
        
        // Date created
        var lCellDateCreated = lModelRow.createCell(8);
        lCellDateCreated.setCellValue(psDateCreated);
        lCellDateCreated.setCellStyle(lDataRowStyle);
       
        // Date changed
        var lCellDateChanged = lModelRow.createCell(9);
        lCellDateChanged.setCellValue(lsDateChanged);
        lCellDateChanged.setCellStyle(lDataRowStyle);
        
        // Document name
        var lCellDocumentName = lModelRow.createCell(10);
        lCellDocumentName.setCellValue("");
        lCellDocumentName.setCellStyle(lDataRowStyle);
        
        // Document links
        var lCellDocumentLinks = lModelRow.createCell(11);
        lCellDocumentLinks.setCellValue(psLinksCellValue);
        lDataRowStyle.setWrapText(true);
        lCellDocumentLinks.setCellStyle(lDataRowStyle);
                
        // up counter
        giRowCounter++;
    }
    catch(ex){
        Dialogs.MsgBox("Error : " + ex );
    }
}

function printSheetRowDocument(lsModelLevel1Name, lsModelLevel2Name, lsModelLevel3Name, lsModelLevel4Name, psDocumentName, psPBO, psDescription, lsBusinessRoles, psDateCreated, lsDateChanged, psLinksCellValue, piLevel, pExcelWorkBook){ 
    try{
        
        var lMainSheet = pExcelWorkBook.getSheetAt(0);
        
        // second row data
        var lModelRow = lMainSheet.createRow(giRowCounter);
        
        var lCell0 = lModelRow.createCell(0);
        lCell0.setCellValue(lsModelLevel1Name);
        lCell0.setCellStyle(lDataRowStyle);
        
        var lCell1 = lModelRow.createCell(1);
        lCell1.setCellValue(lsModelLevel2Name);
        lCell1.setCellStyle(lDataRowStyle);
        
        var lCell2 = lModelRow.createCell(2);
        lCell2.setCellValue(lsModelLevel3Name);
        lCell2.setCellStyle(lDataRowStyle);
        
        var lCell3 = lModelRow.createCell(3);
        lCell3.setCellValue(lsModelLevel4Name);
        lCell3.setCellStyle(lDataRowStyle);        
        
        
        // Activity Name
        var lCellBPO = lModelRow.createCell(4);
        lCellBPO.setCellValue("");
        lCellBPO.setCellStyle(lDataRowStyle);
        
        // BPO Delegate
        var lCellBPO = lModelRow.createCell(5);
        lCellBPO.setCellValue(psPBO);
        lCellBPO.setCellStyle(lDataRowStyle);
        
        // Description
        var lCellDesc = lModelRow.createCell(6);
        lCellDesc.setCellValue(psDescription);
        lCellDesc.setCellStyle(lDataRowStyle);
        
        // Business process role
        var lCellBusinessProcessRole = lModelRow.createCell(7);
        lCellBusinessProcessRole.setCellValue(lsBusinessRoles);
        lCellBusinessProcessRole.setCellStyle(lDataRowStyle);
        
        // Date created
        var lCellDateCreated = lModelRow.createCell(8);
        lCellDateCreated.setCellValue(psDateCreated);
        lCellDateCreated.setCellStyle(lDataRowStyle);
       
        // Date changed
        var lCellDateChanged = lModelRow.createCell(9);
        lCellDateChanged.setCellValue(lsDateChanged);
        lCellDateChanged.setCellStyle(lDataRowStyle);
        
        // Document name
        var lCellDocumentName = lModelRow.createCell(10);
        lCellDocumentName.setCellValue(psDocumentName);
        lCellDocumentName.setCellStyle(lDataRowStyle);
        
        // Document links
        var lCellDocumentLinks = lModelRow.createCell(11);
        lCellDocumentLinks.setCellValue(psLinksCellValue);
        lDataRowStyle.setWrapText(true);
        lCellDocumentLinks.setCellStyle(lDataRowStyle);
                
        // up counter
        giRowCounter++;
    }
    catch(ex){
        Dialogs.MsgBox("Error : " + ex );
    }
}
function businessRolesToString(paBusinessProcessRoles){
    var psResult = "";
    var liCounter = 0;
    for each (loBusinessRole in paBusinessProcessRoles){
        if (liCounter!=0){
            psResult += "\n";
        }
        psResult += loBusinessRole.toString();
        
    }
    return psResult;
}
function getModelLevel(pModel){
    var liModelTypeNum = pModel.OrgModelTypeNum();
    var liResultModelLevel;
    // Process Flow Chart - Standard
    if (liModelTypeNum==giProcessFlowChartStandard){
        liResultModelLevel = 0;
    }
    // Process Flow Chart - SwimLane
    else if (liModelTypeNum==giProcessFlowChartStandardSwimlanes){
        liResultModelLevel = 3;
    }
    return liResultModelLevel;
}
function printLine(){}

// remove line break from String parameter psText
function filterCarriageReturn(psText){
    var lsResult = "";
    lsResult = psText.replaceAll("\\r|\\n", "");
    return lsResult;
}

main()