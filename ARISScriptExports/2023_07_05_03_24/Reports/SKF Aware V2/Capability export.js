var nLocale = Context.getSelectedLanguage()
var oOutput = Context.createOutputObject()
var excelObject = Context.createExcelWorkbook(Context.getSelectedFile());

var selectedDatabases = ArisData.getSelectedDatabases()

//setup excell file
//*************************************************************
var sheetName = "Sheet1"
var mainSheet = excelObject.createSheet(sheetName);
//mainSheet.setColumnWidth(0, 12000);
mainSheet.setDefaultColumnWidth(22)
var rowCounter = 0;
createHeaderRow(mainSheet)
//*************************************************************
//*************************************************************
//Objectives
var sheetName = "Sheet2"
var objectivesSheet = excelObject.createSheet(sheetName);
//mainSheet.setColumnWidth(0, 12000);
objectivesSheet.setDefaultColumnWidth(22)
var rowCounterObjectives = 0;
rowCounterObjectives = createHeaderRowCustom(objectivesSheet, "Objectives", rowCounterObjectives)
//*************************************************************
//Processes
var sheetName = "Sheet3"
var processesSheet = excelObject.createSheet(sheetName);
//mainSheet.setColumnWidth(0, 12000);
processesSheet.setDefaultColumnWidth(22)
var rowCounterProcesses = 0;
rowCounterProcesses = createHeaderRowCustom(processesSheet, "Processes", rowCounterProcesses)
//*************************************************************
//Applications
var sheetName = "Sheet4"
var applicationsSheet = excelObject.createSheet(sheetName);
//mainSheet.setColumnWidth(0, 12000);
applicationsSheet.setDefaultColumnWidth(22)
var rowCounterApplications = 0;
rowCounterApplications = createHeaderRowApplication(applicationsSheet, "Applications", rowCounterApplications)
//*************************************************************
//Semantic data objects
var sheetName = "Sheet5"
var semanticDataObjectsSheet = excelObject.createSheet(sheetName);
//mainSheet.setColumnWidth(0, 12000);
semanticDataObjectsSheet.setDefaultColumnWidth(22)
var rowCounterSemanticDataObjects = 0;
rowCounterSemanticDataObjects = createHeaderRowCustom(semanticDataObjectsSheet, "SemanticDataObjects", rowCounterSemanticDataObjects)
//*************************************************************
//Organization objects
var sheetName = "Sheet6"
var organizationObjectsSheet = excelObject.createSheet(sheetName);
//mainSheet.setColumnWidth(0, 12000);
organizationObjectsSheet.setDefaultColumnWidth(22)
var rowCounterOrganizationObjects = 0;
rowCounterOrganizationObjects = createHeaderRowOrganization(organizationObjectsSheet, "Role", rowCounterOrganizationObjects)
//*************************************************************
//Services/Products
var sheetName = "Sheet7"
var servicesProductsObjectsSheet = excelObject.createSheet(sheetName);
//mainSheet.setColumnWidth(0, 12000);
servicesProductsObjectsSheet.setDefaultColumnWidth(22)
var rowCounterServicesProducts = 0;
rowCounterServicesProducts = createHeaderRowCustom(servicesProductsObjectsSheet, "ServicesProducts", rowCounterServicesProducts)
//*************************************************************
//Task Initiatives
var sheetName = "Sheet8"
var initiativesObjectsSheet = excelObject.createSheet(sheetName);
//mainSheet.setColumnWidth(0, 12000);
initiativesObjectsSheet.setDefaultColumnWidth(22)
var rowCounterInitiatives = 0;
rowCounterInitiatives = createHeaderRowCustom(initiativesObjectsSheet, "Initiatives", rowCounterInitiatives)
//*************************************************************

var selectedDatabase = selectedDatabases[0]
var capabilityObjDefs = selectedDatabase.Find(Constants.SEARCH_OBJDEF, Constants.OT_IS_FUNC)

//var capabilityObjDefs = ArisData.getSelectedObjDefs() //FOR TEST

//FOR TEST
//var selectedGroups = ArisData.getSelectedGroups()
//var capabilityObjDefs = selectedGroups[0].ObjDefList(true, [Constants.OT_IS_FUNC])

for (var i = 0; i < capabilityObjDefs.length; i++) {

    if (capabilityObjDefs[i].Name(nLocale) == "Business Administration")
        var a = "test"

    var capabilityName = capabilityObjDefs[i].Name(nLocale)
    var capabilityGUID = capabilityObjDefs[i].GUID()

    var capabilityLevel = capabilityObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("149d8a20-2def-11e4-7f33-b4b52f5aab7c"), nLocale).getValue()
    var capabilitySource = getCapabilitySource(capabilityObjDefs[i], capabilityLevel)

    var capabilityDescription = capabilityObjDefs[i].Attribute(Constants.AT_DESC, nLocale).getValue()

    var asIsMaturity = capabilityObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("bb9a31a0-8901-11ec-1a81-0022487eead8"), nLocale).getValue()
    var targetMaturity = capabilityObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("d61b03b1-8901-11ec-1a81-0022487eead8"), nLocale).getValue()
    var releaseLastReview = capabilityObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("f2c19bf0-8901-11ec-1a81-0022487eead8"), nLocale).getValue()

    var referenceCapabilityMaps = capabilityObjDefs[i].getConnectedObjs([Constants.OT_IS_FUNC], Constants.EDGES_INOUT, [Constants.CT_REFS_TO_3])
    var referenceCapabilityMapName = ""
    if (referenceCapabilityMaps.length > 0)
        var referenceCapabilityMapName = referenceCapabilityMaps[0].Name(nLocale)

    //Strategic Relevance attributes
    var leveragePotentialAsIsCompliance = capabilityObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("d9bc0080-89b7-11ec-1a81-0022487eead8"), nLocale).getValue()
    var leveragePotentialAsIsDifferentiating = capabilityObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("14828a90-89b8-11ec-1a81-0022487eead8"), nLocale).getValue()
    var leveragePotentialAsIsMissionCritical = capabilityObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("b254b8c0-89b7-11ec-1a81-0022487eead8"), nLocale).getValue()
    var leveragePotentialAsIsScale = capabilityObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("facc3471-89b7-11ec-1a81-0022487eead8"), nLocale).getValue()
    var leveragePotentialComment = capabilityObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("32342a31-89b8-11ec-1a81-0022487eead8"), nLocale).getValue()
    var leveragePotentialTrendCompliance = capabilityObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("6b6d5d81-89b8-11ec-1a81-0022487eead8"), nLocale).getValue()
    var leveragePotentialTrendDifferentiating = capabilityObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("a6509751-89b8-11ec-1a81-0022487eead8"), nLocale).getValue()
    var leveragePotentialTrendMissionCritical = capabilityObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("4f371340-89b8-11ec-1a81-0022487eead8"), nLocale).getValue()
    var leveragePotentialTrendScale = capabilityObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("8ad6bf40-89b8-11ec-1a81-0022487eead8"), nLocale).getValue()

    var capabilityReachAsIsGlobal = capabilityObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("df951541-89b8-11ec-1a81-0022487eead8"), nLocale).getValue()
    var capabilityReachAsIsLocal = capabilityObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("c65ec490-89b8-11ec-1a81-0022487eead8"), nLocale).getValue()
    var capabilityReachAsIsProductLine = capabilityObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("f9b32de0-89b8-11ec-1a81-0022487eead8"), nLocale).getValue()
    var capabilityReachAsIsRegional = capabilityObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("11735951-89b9-11ec-1a81-0022487eead8"), nLocale).getValue()
    var capabilityReachComment = capabilityObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("1fed8e60-89b9-11ec-1a81-0022487eead8"), nLocale).getValue()
    var capabilityReachTrendLocal = capabilityObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("33bf4191-89b9-11ec-1a81-0022487eead8"), nLocale).getValue()
    var capabilityReachTrendRegional = capabilityObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("734612d0-89b9-11ec-1a81-0022487eead8"), nLocale).getValue()
    var capabilityReachTrendProductLine = capabilityObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("5b11c8d1-89b9-11ec-1a81-0022487eead8"), nLocale).getValue()
    var capabilityReachTrendGlobal = capabilityObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("477ad4b0-89b9-11ec-1a81-0022487eead8"), nLocale).getValue()


    var processesObjDefs = capabilityObjDefs[i].getConnectedObjs([Constants.OT_FUNC], Constants.EDGES_INOUT, [Constants.CT_CAN_SUPP_1, Constants.CT_SUPP_3, Constants.CT_SUPPORTS_1])
    rowCounterProcesses = objDefListToSheet(processesObjDefs, capabilityGUID, processesSheet, rowCounterProcesses)

    var objectivesObjDefs = capabilityObjDefs[i].getConnectedObjs([Constants.OT_OBJECTIVE], Constants.EDGES_INOUT, [Constants.CT_CAN_SUPP_1, Constants.CT_SUPP_3, Constants.CT_SUPPORTS_1])
    rowCounterObjectives = objDefListToSheet(objectivesObjDefs, capabilityGUID, objectivesSheet, rowCounterObjectives)

    var semanticDataObjectsObjDefs = capabilityObjDefs[i].getConnectedObjs([Constants.OT_ENT_TYPE], Constants.EDGES_INOUT, [Constants.CT_IS_INP_FOR])
    rowCounterSemanticDataObjects = objDefListToSheet(semanticDataObjectsObjDefs, capabilityGUID, semanticDataObjectsSheet, rowCounterSemanticDataObjects)

    var applicationsObjDefs = capabilityObjDefs[i].getConnectedObjs([Constants.OT_APPL_SYS_CLS], Constants.EDGES_INOUT, [Constants.CT_IS_ALLOW_TO_ASSIG_INST, Constants.CT_PROVIDES])
    rowCounterApplications = appObjDefListToSheet(applicationsObjDefs, capabilityGUID, applicationsSheet, rowCounterApplications)

    var organizationObjDefs = capabilityObjDefs[i].getConnectedObjs([Constants.OT_PERS_TYPE], Constants.EDGES_INOUT, [Constants.CT_CAN_BE_USER])
    rowCounterOrganizationObjects = orgObjDefListToSheet(organizationObjDefs, capabilityGUID, organizationObjectsSheet, rowCounterOrganizationObjects)

    var servicesProductsObjDefs = capabilityObjDefs[i].getConnectedObjs([Constants.OT_PERF], Constants.EDGES_INOUT, [Constants.CT_INFLUENCES])
    rowCounterServicesProducts = objDefListToSheet(servicesProductsObjDefs, capabilityGUID, servicesProductsObjectsSheet, rowCounterServicesProducts)

    var initiativesObjDefs = capabilityObjDefs[i].getConnectedObjs([Constants.OT_FUNC_INST], Constants.EDGES_INOUT, [Constants.CT_REFS_TO_2])
    rowCounterInitiatives = objDefListToSheet(initiativesObjDefs, capabilityGUID, initiativesObjectsSheet, rowCounterInitiatives)

    //***********************************************************************************
    //PRINT DATA

    var newRow = mainSheet.createRow(rowCounter);
    rowCounter++

    var newCell = newRow.createCell(0);
    newCell.setCellValue(capabilityGUID);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(1);
    newCell.setCellValue(capabilitySource);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(2);
    newCell.setCellValue(capabilityName);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(3);
    newCell.setCellValue(capabilityLevel);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(4);
    newCell.setCellValue(capabilityDescription);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(5);
    newCell.setCellValue("");
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(6);
    newCell.setCellValue(asIsMaturity);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(7);
    newCell.setCellValue(targetMaturity);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(8);
    newCell.setCellValue(releaseLastReview);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(9);
    newCell.setCellValue(referenceCapabilityMapName);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(10);
    newCell.setCellValue(leveragePotentialAsIsCompliance);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(11);
    newCell.setCellValue(leveragePotentialAsIsDifferentiating);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(12);
    newCell.setCellValue(leveragePotentialAsIsMissionCritical);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(13);
    newCell.setCellValue(leveragePotentialAsIsScale);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(14);
    newCell.setCellValue(leveragePotentialComment);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(15);
    newCell.setCellValue(leveragePotentialTrendCompliance);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(16);
    newCell.setCellValue(leveragePotentialTrendDifferentiating);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(17);
    newCell.setCellValue(leveragePotentialTrendMissionCritical);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(18);
    newCell.setCellValue(leveragePotentialTrendScale);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(19);
    newCell.setCellValue(capabilityReachAsIsGlobal);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(20);
    newCell.setCellValue(capabilityReachAsIsLocal);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(21);
    newCell.setCellValue(capabilityReachAsIsProductLine);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(22);
    newCell.setCellValue(capabilityReachAsIsRegional);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(23);
    newCell.setCellValue(capabilityReachComment);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(24);
    newCell.setCellValue(capabilityReachTrendLocal);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(25);
    newCell.setCellValue(capabilityReachTrendRegional);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(26);
    newCell.setCellValue(capabilityReachTrendProductLine);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(27);
    newCell.setCellValue(capabilityReachTrendGlobal);
    //newCell.setCellStyle(cellStyle);

    //Strategic Relevance attributes
    //var leveragePotentialAsIsCompliance
    //var leveragePotentialAsIsDifferentiating 
    //var leveragePotentialAsIsMissionCritical
    //var leveragePotentialAsIsScale
    //var leveragePotentialComment 
    //var leveragePotentialTrendCompliance
    //var leveragePotentialTrendDifferentiating 
    //var leveragePotentialTrendMissionCritical 
    //var leveragePotentialTrendScale 
    //var capabilityReachAsIsGlobal
    //var capabilityReachAsIsLocal 
    //var capabilityReachAsIsProductLine 
    //var capabilityReachAsIsRegional 
    //var capabilityReachComment
    //var capabilityReachTrendLocal 
    //var capabilityReachTrendRegional
    //var capabilityReachTrendProductLine 
    //var capabilityReachTrendGlobal 

    //***********************************************************************************
}


excelObject.write()
writeDocumentADS(excelObject) //copy document to the ads

function writeDocumentADS(excelObject) {
    var adsReport = Context.getComponent("ADS")
    var repository = adsReport.getADSRepository("portal")
    var rootFolder = repository.getRootFolder();

    var folderName = "SKF Capability Dashboard\\"
    var adsFolder = repository.getFolder(rootFolder, folderName);

    if (!adsFolder)
        adsFolder = repository.createFolder(rootFolder, folderName);

    var docInfo = repository.createDocumentMetaInfo("Capability data", "capabilityData.xlsx", "Capability data");

    //contentByteArray = excelObject.getDocumentContent()

    var contentByteArray = Context.getFile(Context.getSelectedFile(), Constants.LOCATION_OUTPUT)

    var document = repository.createAndOverwriteExistingDocument(adsFolder, docInfo, new java.io.ByteArrayInputStream(contentByteArray))
}

function getCapabilitySource(capabilityObjDef, level) {

    if (level == "Level 1") {
        //get source from connected object
        var tempObjDefs = capabilityObjDef.getConnectedObjs([Constants.OT_IS_FUNC], Constants.EDGES_IN, [Constants.CT_CALLS_1])
        var tempGUID = null
        if (tempObjDefs.length > 0) {
            var tempGUID = tempObjDefs[0].GUID()
        }

        if (tempGUID != null) {
            return tempGUID
        }
        else {
            return "ERROR! Parent object not found"
        }
    }

    else if (level == "Level 2") {
    
        var tempObjDefs = capabilityObjDef.getConnectedObjs([Constants.OT_IS_FUNC], Constants.EDGES_IN, [Constants.CT_CALLS_1])
        var tempGUID = null

        for(var i=0; i<tempObjDefs.length; i++){
            var capabilityLevelTemp = tempObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("149d8a20-2def-11e4-7f33-b4b52f5aab7c"), nLocale).getValue()
            if(capabilityLevelTemp = "Level 1"){
            var tempGUID = tempObjDefs[0].GUID()
            }
        }
        
        if (tempGUID != null) {
            return tempGUID
        }
        else {
            return "ERROR! Parent object not found"
        }
        
        
/*         //for test
        //if(capabilityObjDef.Name(nLocale)=="Manage Delivery")
        //var a = "STOP"

        var modelTemp = null
        var occListTemp = capabilityObjDef.OccList()
        for (var i = 0; i < occListTemp.length; i++) {
            if (occListTemp[i].Model().Type() == "Capability Map") {
                var modelTemp = occListTemp[i].Model()
            }
        }
        //get model where the object is occuring, from that model get object where the assigment is created
        //var modelsTemp = capabilityObjDef.AssignedModels(ArisData.ActiveFilter().UserDefinedModelTypeNum("1da34d91-9935-11eb-6971-00505616050e")) //get capability model where this object is located
        //var occListTemp = capabilityObjDef.OccList([ArisData.ActiveFilter().UserDefinedModelTypeNum("1da34d91-9935-11eb-6971-00505616050e")]) //get all occurences from the capability map model 

        var tempGUID = null;
        if (modelTemp != null) {
            //var modelTemp = modelsTemp[0]
            var assignedObjects = modelTemp.getSuperiorObjDefs([Constants.OT_IS_FUNC]) //get superior objects assigned filter this only get level 1

            for (var i = 0; i < assignedObjects.length; i++) {
                var capabilityLevelTemp = assignedObjects[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("149d8a20-2def-11e4-7f33-b4b52f5aab7c"), nLocale).getValue()

                if (capabilityLevelTemp == "Level 1") {
                    //return assignedObjects[i].GUID() //this is parent
                    var tempGUID = assignedObjects[i].GUID()
                }
            }
        }
        if (tempGUID != null) {
            return tempGUID
        }
        else {
            return "ERROR! Parent object not found"
        } */
    }

    else if (level == "Level 3") {
        //get source from connected object
        var tempObjDefs = capabilityObjDef.getConnectedObjs([Constants.OT_IS_FUNC], Constants.EDGES_IN, [Constants.CT_CALLS_1])
        var tempGUID = null
        if (tempObjDefs.length > 0) {
            var tempGUID = tempObjDefs[0].GUID()
        }
        if (tempGUID != null) {
            return tempGUID
        }
        else {
            return "ERROR! Parent object not found"
        }
    }

    else if (level == "Level 0") {
        return "";
    }

    else {
        return "ERROR! Parent object not found"
    }
}

function createHeaderRow(headerRow) {
    var headerRow = mainSheet.createRow(rowCounter);
    rowCounter++;
    var font = excelObject.createFont();
    cellStyle1 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(169, 208, 142), Constants.SOLID_FOREGROUND);
    cellStyle2 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(155, 194, 230), Constants.SOLID_FOREGROUND);
    cellStyle2.setLocked(true)
    cellStyle3 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(255, 217, 102), Constants.SOLID_FOREGROUND);

    var newCell = headerRow.createCell(0);
    newCell.setCellValue("GUID");
    newCell.setCellStyle(cellStyle1);

    var newCell = headerRow.createCell(1);
    newCell.setCellValue("Parent GUID");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(2);
    newCell.setCellValue("Name");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(3);
    newCell.setCellValue("Level");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(4);
    newCell.setCellValue("Description");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(5);
    newCell.setCellValue("Processes");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(6);
    newCell.setCellValue("AsIs Maturity");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(7);
    newCell.setCellValue("Target Maturity");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(8);
    newCell.setCellValue("ReleaseLastReview");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(9);
    newCell.setCellValue("referenceCapabilityMap");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(10);
    newCell.setCellValue("leveragePotentialAsIsCompliance");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(11);
    newCell.setCellValue("leveragePotentialAsIsDifferentiating");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(12);
    newCell.setCellValue("leveragePotentialAsIsMissionCritical");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(13);
    newCell.setCellValue("leveragePotentialAsIsScale");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(14);
    newCell.setCellValue("leveragePotentialComment");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(15);
    newCell.setCellValue("leveragePotentialTrendCompliance");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(16);
    newCell.setCellValue("leveragePotentialTrendDifferentiating");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(17);
    newCell.setCellValue("leveragePotentialTrendMissionCritical");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(18);
    newCell.setCellValue("leveragePotentialTrendScale");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(19);
    newCell.setCellValue("capabilityReachAsIsGlobal");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(20);
    newCell.setCellValue("capabilityReachAsIsLocal");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(21);
    newCell.setCellValue("capabilityReachAsIsProductLine");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(22);
    newCell.setCellValue("capabilityReachAsIsRegional");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(23);
    newCell.setCellValue("capabilityReachComment");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(24);
    newCell.setCellValue("capabilityReachTrendLocal");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(25);
    newCell.setCellValue("capabilityReachTrendRegional");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(26);
    newCell.setCellValue("capabilityReachTrendProductLine");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(27);
    newCell.setCellValue("capabilityReachTrendGlobal");
    newCell.setCellStyle(cellStyle2);
}

function createHeaderRowCustom(sheet, columnName, rowCn) {
    var headerRow = sheet.createRow(rowCn);
    rowCn++;
    var font = excelObject.createFont();
    cellStyle1 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(169, 208, 142), Constants.SOLID_FOREGROUND);
    cellStyle2 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(155, 194, 230), Constants.SOLID_FOREGROUND);
    cellStyle2.setLocked(true)
    cellStyle3 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(255, 217, 102), Constants.SOLID_FOREGROUND);

    var newCell = headerRow.createCell(0);
    newCell.setCellValue("GUID");
    newCell.setCellStyle(cellStyle1);

    var newCell = headerRow.createCell(1);
    newCell.setCellValue(columnName);
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(2);
    newCell.setCellValue("Link");
    newCell.setCellStyle(cellStyle1);

    return rowCn;
}

function RGB(r, g, b) {
    return (new java.awt.Color(r / 255.0, g / 255.0, b / 255.0, 1)).getRGB() & 0xFFFFFF
}

function objDefListToSheet(objDefs, oGuid, sheet, rowCn) {
    for (var i = 0; i < objDefs.length; i++) {
        var value = objDefs[i].Name(nLocale)
        var objectLink = getPortalLink(objDefs[i])

        var newRow = sheet.createRow(rowCn);
        rowCn++

        var newCell = newRow.createCell(0);
        newCell.setCellValue(oGuid);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(1);
        newCell.setCellValue(value);
        //newCell.setCellStyle(cellStyle);   

        var newCell = newRow.createCell(2);
        newCell.setCellValue(objectLink);
        //newCell.setCellStyle(cellStyle);   
    }
    return rowCn;
}

function appObjDefListToSheet(objDefs, oGuid, sheet, rowCn) {
    for (var i = 0; i < objDefs.length; i++) {
        var appName = objDefs[i].Name(nLocale)
        var appLink = getPortalLink(objDefs[i])

        var applicationUsage = objDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("15e4af10-7853-11ec-1a81-0022487eead8"), nLocale).getValue()

        var CMOtechnicalLifeCycleStatus = objDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("afdeb5a0-7e04-11ec-1a81-0022487eead8"), nLocale).getValue()
        var FMOtechnicalLifeCycleStatus = objDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("42b13a10-7e05-11ec-1a81-0022487eead8"), nLocale).getValue()
        var IMOtechnicalLifeCycleStatus = objDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("fc6725b0-7e04-11ec-1a81-0022487eead8"), nLocale).getValue()

        var CMOdeploymentMode = objDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("aa247380-7e03-11ec-1a81-0022487eead8"), nLocale).getValue()
        var FMOdeploymentMode = objDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("4f829690-7e04-11ec-1a81-0022487eead8"), nLocale).getValue()
        var IMOdeploymentMode = objDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("01e7a510-7e04-11ec-1a81-0022487eead8"), nLocale).getValue()

        var newRow = sheet.createRow(rowCn);
        rowCn++

        var newCell = newRow.createCell(0);
        newCell.setCellValue(oGuid);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(1);
        newCell.setCellValue(appName);
        //newCell.setCellStyle(cellStyle);   

        var newCell = newRow.createCell(2);
        newCell.setCellValue(appLink);
        //newCell.setCellStyle(cellStyle);  

        var newCell = newRow.createCell(3);
        newCell.setCellValue(applicationUsage);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(4);
        newCell.setCellValue(CMOtechnicalLifeCycleStatus);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(5);
        newCell.setCellValue(FMOtechnicalLifeCycleStatus);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(6);
        newCell.setCellValue(IMOtechnicalLifeCycleStatus);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(7);
        newCell.setCellValue(CMOdeploymentMode);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(8);
        newCell.setCellValue(FMOdeploymentMode);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(9);
        newCell.setCellValue(IMOdeploymentMode);
        //newCell.setCellStyle(cellStyle);
    }
    return rowCn;
}

function orgObjDefListToSheet(objDefs, oGuid, sheet, rowCn) {
    for (var i = 0; i < objDefs.length; i++) {
        var roleName = objDefs[i].Name(nLocale)
        var roleLink = getPortalLink(objDefs[i])

        var organizationType = objDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("fcb3c040-8a53-11ec-1a81-0022487eead8"), nLocale).getValue()

        var connectedPersons = objDefs[i].getConnectedObjs([Constants.OT_PERS], Constants.EDGES_INOUT, [Constants.CT_EXEC_1])
        var connectedPersonName = connectedPersons[0].Name(nLocale)
        var connectedPersonEmail = connectedPersons[0].Attribute(Constants.AT_EMAIL_ADDR, nLocale).getValue()

        var newRow = sheet.createRow(rowCn);
        rowCn++

        var newCell = newRow.createCell(0);
        newCell.setCellValue(oGuid);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(1);
        newCell.setCellValue(roleName);
        //newCell.setCellStyle(cellStyle);   

        var newCell = newRow.createCell(2);
        newCell.setCellValue(organizationType);
        //newCell.setCellStyle(cellStyle);  

        var newCell = newRow.createCell(3);
        newCell.setCellValue(connectedPersonName);
        //newCell.setCellStyle(cellStyle);  

        var newCell = newRow.createCell(4);
        newCell.setCellValue(connectedPersonEmail);
        //newCell.setCellStyle(cellStyle);

    }
    return rowCn;
}

function createHeaderRowOrganization(sheet, columnName, rowCn) {
    var headerRow = sheet.createRow(rowCn);
    rowCn++;
    var font = excelObject.createFont();
    cellStyle1 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(169, 208, 142), Constants.SOLID_FOREGROUND);
    cellStyle2 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(155, 194, 230), Constants.SOLID_FOREGROUND);
    cellStyle2.setLocked(true)
    cellStyle3 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(255, 217, 102), Constants.SOLID_FOREGROUND);

    var newCell = headerRow.createCell(0);
    newCell.setCellValue("GUID");
    newCell.setCellStyle(cellStyle1);

    var newCell = headerRow.createCell(1);
    newCell.setCellValue(columnName);
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(2);
    newCell.setCellValue("Type");
    newCell.setCellStyle(cellStyle1);

    var newCell = headerRow.createCell(3);
    newCell.setCellValue("Person name");
    newCell.setCellStyle(cellStyle1);

    var newCell = headerRow.createCell(4);
    newCell.setCellValue("Person email");
    newCell.setCellStyle(cellStyle1);

    return rowCn;
}

function createHeaderRowApplication(sheet, columnName, rowCn) {
    var headerRow = sheet.createRow(rowCn);
    rowCn++;
    var font = excelObject.createFont();
    cellStyle1 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(169, 208, 142), Constants.SOLID_FOREGROUND);
    cellStyle2 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(155, 194, 230), Constants.SOLID_FOREGROUND);
    cellStyle2.setLocked(true)
    cellStyle3 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(255, 217, 102), Constants.SOLID_FOREGROUND);

    var newCell = headerRow.createCell(0);
    newCell.setCellValue("GUID");
    newCell.setCellStyle(cellStyle1);

    var newCell = headerRow.createCell(1);
    newCell.setCellValue(columnName);
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(2);
    newCell.setCellValue("Link");
    newCell.setCellStyle(cellStyle1);

    var newCell = headerRow.createCell(3);
    newCell.setCellValue("applicationUsage");
    newCell.setCellStyle(cellStyle1);

    var newCell = headerRow.createCell(4);
    newCell.setCellValue("CMOtechnicalLifeCycleStatus");
    newCell.setCellStyle(cellStyle1);

    var newCell = headerRow.createCell(5);
    newCell.setCellValue("FMOtechnicalLifeCycleStatus");
    newCell.setCellStyle(cellStyle1);

    var newCell = headerRow.createCell(6);
    newCell.setCellValue("IMOtechnicalLifeCycleStatus");
    newCell.setCellStyle(cellStyle1);

    var newCell = headerRow.createCell(7);
    newCell.setCellValue("CMOdeploymentMode");
    newCell.setCellStyle(cellStyle1);

    var newCell = headerRow.createCell(8);
    newCell.setCellValue("FMOdeploymentMode");
    newCell.setCellStyle(cellStyle1);

    var newCell = headerRow.createCell(9);
    newCell.setCellValue("IMOdeploymentMode");
    newCell.setCellStyle(cellStyle1);

    return rowCn;
}

function getPortalLink(selectedObject) {
    //http://<servername>/#<tenant>/item/c.<factSheet>.<DBName>.<GUID>.<Version>
    var oDatabase = ArisData.getActiveDatabase();
    var sTenantName = ArisData.getTenantName();
    var serverName = oDatabase.ServerName();
    var sDBName = oDatabase.Name(nLocale);
    sDBName = sDBName.replace(" ", "%20");
    var mGUID = selectedObject.GUID()

    var linkURL = "https://" + serverName + "/#" + sTenantName + "/item/c." + sDBName + "." + mGUID + ".-1"
    return linkURL
}