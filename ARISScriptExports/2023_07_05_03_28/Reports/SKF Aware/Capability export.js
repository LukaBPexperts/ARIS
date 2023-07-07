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
rowCounterApplications = createHeaderRowCustom(applicationsSheet, "Applications", rowCounterApplications)
//*************************************************************
//Semantic data objects
var sheetName = "Sheet5"
var semanticDataObjectsSheet = excelObject.createSheet(sheetName);
//mainSheet.setColumnWidth(0, 12000);
semanticDataObjectsSheet.setDefaultColumnWidth(22)
var rowCounterSemanticDataObjects = 0;
rowCounterSemanticDataObjects = createHeaderRowCustom(semanticDataObjectsSheet, "SemanticDataObjects", rowCounterSemanticDataObjects)
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

    var processesObjDefs = capabilityObjDefs[i].getConnectedObjs([Constants.OT_FUNC], Constants.EDGES_INOUT, [Constants.CT_CAN_SUPP_1, Constants.CT_SUPP_3, Constants.CT_SUPPORTS_1])
    rowCounterProcesses = objDefListToSheet(processesObjDefs, capabilityGUID, processesSheet, rowCounterProcesses)

    var objectivesObjDefs = capabilityObjDefs[i].getConnectedObjs([Constants.OT_OBJECTIVE], Constants.EDGES_INOUT, [Constants.CT_CAN_SUPP_1, Constants.CT_SUPP_3, Constants.CT_SUPPORTS_1])
    rowCounterObjectives = objDefListToSheet(objectivesObjDefs, capabilityGUID, objectivesSheet, rowCounterObjectives)

    var semanticDataObjectsObjDefs = capabilityObjDefs[i].getConnectedObjs([Constants.OT_ENT_TYPE], Constants.EDGES_INOUT, [Constants.CT_IS_INP_FOR])
    rowCounterSemanticDataObjects = objDefListToSheet(semanticDataObjectsObjDefs, capabilityGUID, semanticDataObjectsSheet, rowCounterSemanticDataObjects)

    var applicationsObjDefs = capabilityObjDefs[i].getConnectedObjs([Constants.OT_APPL_SYS_CLS], Constants.EDGES_INOUT, [Constants.CT_IS_ALLOW_TO_ASSIG_INST, Constants.CT_PROVIDES])
    rowCounterApplications = objDefListToSheet(applicationsObjDefs, capabilityGUID, applicationsSheet, rowCounterApplications)

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

        //for test
        //if(capabilityObjDef.Name(nLocale)=="Manage Audit & Internal Control")
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
        }
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