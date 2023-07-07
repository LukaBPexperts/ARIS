var nLocale = Context.getSelectedLanguage()
var oOutput = Context.createOutputObject()
var selectedDatabases = ArisData.getSelectedDatabases()

//Legal Entity group path
var legalEntitiesGrpPath = "Main group/Legal Entities" 

//User defined attribute GUIDs
var g_modelType = "3bb28c60-1b30-11ec-6580-005056bd3657"
var g_e2eDomain = "31338450-1b31-11ec-6580-005056bd3657"

//User defined symbol GUIDs
var legalEntitySymbol = "4589e5f0-43ab-11e2-28c9-0050568e0118" //Legal entity view	

//Possible Template/Divison Values 
var templateValuesList = ["MGF", "RISE", "LEAN", "NEXT", "NEO", "CTS", "Commercial"];

//setup excell file
//*************************************************************
var byteArrayFile = readDocumentADS() //get file from ADS
if(!byteArrayFile){
    //excel not available in ADS
    var excelObject = Context.createExcelWorkbook(Context.getSelectedFile());
    
    var sheetName_0 = "Legal Entities"
    var sheetName_1 = "Operating Models"
    var sheetName_2 = "End to End Processes"
    var sheetName_3 = "Process Owners"
    var sheetName_4 = "Sector Templates"
    var sheetName_5 = "Description"
    
    var sheet_0 = excelObject.createSheet(sheetName_0);
    var sheet_1 = excelObject.createSheet(sheetName_1);
    var sheet_2 = excelObject.createSheet(sheetName_2);
    var sheet_3 = excelObject.createSheet(sheetName_3);
    var sheet_4 = excelObject.createSheet(sheetName_4);
    var sheet_5 = excelObject.createSheet(sheetName_5);

}
else{
    //excel exist in ADS
    var excelObject = Context.createExcelWorkbook(Context.getSelectedFile(), byteArrayFile)
    
    var sheetName_0 = "Legal Entities"
    var sheetName_1 = "Operating Models"
    var sheetName_2 = "End to End Processes"
    var sheetName_3 = "Process Owners"
    var sheetName_4 = "Sector Templates"
    //var sheetName_5 = "Description"

    var sheet_0 = getSheetByName(sheetName_0, excelObject)
    var sheet_1 = getSheetByName(sheetName_1, excelObject)
    var sheet_2 = getSheetByName(sheetName_2, excelObject)
    var sheet_3 = getSheetByName(sheetName_3, excelObject)
    var sheet_4 = getSheetByName(sheetName_4, excelObject)
    //var sheet_5 = getSheetByName(sheetName_5, excelObject)
    
    //remove existing data 
    removeSheetContent(sheet_0)
    removeSheetContent(sheet_1)
    removeSheetContent(sheet_2)
    removeSheetContent(sheet_3)
    removeSheetContent(sheet_4)
    //removeSheetContent(sheet_5)
}

var rowCounter_0 = 0;
var rowCounter_1 = 0;
var rowCounter_2 = 0;
var rowCounter_3 = 0;
var rowCounter_4 = 0;

createHeaderRowSheet_0(sheet_0) //Legal Entities
createHeaderRowSheet_1(sheet_1) //Operating Models
createHeaderRowSheet_2(sheet_2) //End to End Processes
createHeaderRowSheet_3(sheet_3) //Process Owners
createHeaderRowSheet_4(sheet_4) //Sector Templates

//*************************************************************

var selectedDatabase = selectedDatabases[0]

//*************************************************************
//***Sheet Legal Entities***
//*************************************************************
var legalEntitiesGrp = selectedDatabase.Group(legalEntitiesGrpPath, nLocale)
if(legalEntitiesGrp.IsValid()){
    //get Legal Entities data
    var continetModels = legalEntitiesGrp.ModelListFilter(Constants.MT_VAL_ADD_CHN_DGM)
    //for each cont model get func, from each func assigned vacd model, from each assigned vacd model Legal Entities objects
    for (var i = 0; i < continetModels.length; i++) {
        var countryObjDefs = continetModels[i].ObjDefListByTypes([Constants.OT_FUNC])
        for (var j = 0; j < countryObjDefs.length; j++) {
            var countryModel = countryObjDefs[j].AssignedModels(Constants.MT_VAL_ADD_CHN_DGM)[0]
            var legalEntityObjDefs = new Array()
            if(countryModel != null){
                var legalEntityObjDefs = countryModel.ObjDefListBySymbols([ArisData.ActiveFilter().UserDefinedSymbolTypeNum(legalEntitySymbol)])
            }
            for (var k = 0; k < legalEntityObjDefs.length; k++) {
                var fullNameLegalEntity = legalEntityObjDefs[k].Attribute(Constants.AT_NAME_FULL, nLocale).getValue()
                print_0(continetModels[i].Name(nLocale), getPortalLink(continetModels[i]), countryObjDefs[j].Name(nLocale), getPortalLink(countryModel), legalEntityObjDefs[k].Name(nLocale), fullNameLegalEntity, sheet_0)
            }
        }
    }
}
//*************************************************************

//*************************************************************
//***Sheet Operating Models, End to End Processes***
//*************************************************************
var modelsList = selectedDatabase.Find(Constants.SEARCH_MODEL)

for (var i = 0; i < modelsList.length; i++) {
    var model = modelsList[i]
    var modelName = model.Name(nLocale)
    var modelType = model.Type()
    var modelLink = getPortalLink(model)
    
    var processOwnerAtt = model.Attribute(Constants.AT_PERS_RESP, nLocale).getValue()
    var e2eDomainAtt = model.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum(g_e2eDomain), nLocale).getValue()
    var modelTypeAtt = model.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum(g_modelType), nLocale).getValue()
    
    //get template name from folder structure
    var templateName = ""
    var domainName = ""
    var grpFlag = false
    var modelGrp = model.Group()    
    if(modelGrp.Name(nLocale).toLowerCase().includes("main processes")){
        var grpFlag = true
        var domainGrp = modelGrp.Parent()        
        for (var j = 0; j < templateValuesList.length; j++) {
            //get template Name
            if (modelGrp.Name(nLocale).toLowerCase().startsWith(templateValuesList[j].toLowerCase())) {
            var templateName = templateValuesList[j];
            //break;
            }
            //get Domain name
            if (domainGrp.Name(nLocale).toLowerCase().startsWith(templateValuesList[j].toLowerCase())) {
                var domainName = domainGrp.Name(nLocale).substring(templateValuesList[j].length +1)
            }      
        }
    }
    
    if(modelTypeAtt.equals("OPM - Operating model")){               
        print_1(modelName, modelLink, sheet_1) //fill Opearting Models sheet      
    }
    else if(modelTypeAtt.equals("Main process")){       
        print_2(modelType, modelName, modelLink, sheet_2) //fill End to End sheet
        print_3(processOwnerAtt, e2eDomainAtt, modelName, modelLink, modelTypeAtt, sheet_3) //fill Process Owners sheet
         
        //check if model is in Main process group, if not print out empty att
        if(grpFlag){        
            print_4(templateName, domainName, modelName, modelLink, sheet_4) //fill Sector Templates
        }
        else{
            print_4("", "", modelName, modelLink, sheet_4) //fill Sector Templates
        }
    } 
    else if(modelTypeAtt.equals("Subprocess") || modelTypeAtt.equals("Variants model") || modelTypeAtt.equals("FAD")){              
        print_3(processOwnerAtt, e2eDomainAtt, modelName, modelLink, modelTypeAtt, sheet_3) //fill Process Owners sheet
    } 
}
//*************************************************************

excelObject.write()
writeDocumentADS(excelObject) //write document to the ADS

function writeDocumentADS(excelObject) {
    var adsReport = Context.getComponent("ADS")
    var repository = adsReport.getADSRepository("portal")
    var rootFolder = repository.getRootFolder();

    var mercFolderName = "Merck Entry Page\\"
    var adsMercFolder = repository.getFolder(rootFolder, mercFolderName);
    if (!adsMercFolder)
        adsMercFolder = repository.createFolder(rootFolder, mercFolderName);
    
    var dataFolderName = "data\\"
    var adsDataFolder = repository.getFolder(adsMercFolder, dataFolderName);
    if (!adsDataFolder)
        adsDataFolder = repository.createFolder(adsMercFolder, dataFolderName);
    

    var docInfo = repository.createDocumentMetaInfo("BPMN_2_input_data_entry-page", "BPMN_2_input_data_entry-page.xlsx", "BPMN_2_input_data_entry-page");
    //contentByteArray = excelObject.getDocumentContent()
    var contentByteArray = Context.getFile(Context.getSelectedFile(), Constants.LOCATION_OUTPUT)
    var document = repository.createAndOverwriteExistingDocument(adsDataFolder, docInfo, new java.io.ByteArrayInputStream(contentByteArray))
}

function readDocumentADS() {
    var adsReport = Context.getComponent("ADS")
    var repository = adsReport.getADSRepository("portal")
    var rootFolder = repository.getRootFolder();

    var mercFolderName = "Merck Entry Page\\"
    var adsMercFolder = repository.getFolder(rootFolder, mercFolderName);
    if (!adsMercFolder){
        //error
        return null;
    }
    
    var dataFolderName = "data\\"
    var adsDataFolder = repository.getFolder(adsMercFolder, dataFolderName);
    if (!adsDataFolder){
        //error
        return null;
    }
    
    var fileName = "BPMN_2_input_data_entry-page.xlsx"
    var documentTemplate = repository.getDocument(adsDataFolder, null, fileName)
    
    if(!documentTemplate){
        //error
        return null;
    }
    else{
        var docContent = documentTemplate.getDocumentContent()
        var byteDoc = Packages.com.google.common.io.ByteStreams.toByteArray(docContent) //convert InputStream to Byte Array
        return byteDoc;
    }
}

function createHeaderRowSheet_0(sheet) {
    var headerRow = sheet.createRow(rowCounter_0);
    rowCounter_0++;
    
    sheet.setColumnWidth(0, 10000);
    sheet.setColumnWidth(1, 10000);
    sheet.setColumnWidth(2, 10000);
    sheet.setColumnWidth(3, 10000);
    
    var font = excelObject.createFont();
    font.setColor(new java.awt.Color(255 / 255.0, 255 / 255.0, 255 / 255.0, 1))
    font.setBold(true)
    
    cellStyle1 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(68, 114, 196), Constants.SOLID_FOREGROUND);
    //cellStyle2 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(155, 194, 230), Constants.SOLID_FOREGROUND);
    //cellStyle2.setLocked(true)
    //cellStyle3 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(255, 217, 102), Constants.SOLID_FOREGROUND);

    var newCell = headerRow.createCell(0);
    newCell.setCellValue("Region of Merck");
    newCell.setCellStyle(cellStyle1);

    var newCell = headerRow.createCell(1);
    newCell.setCellValue("Country");
    newCell.setCellStyle(cellStyle1);

    var newCell = headerRow.createCell(2);
    newCell.setCellValue("Legal Entity");
    newCell.setCellStyle(cellStyle1);
    
    var newCell = headerRow.createCell(3);
    newCell.setCellValue("Legal Entity Name");
    newCell.setCellStyle(cellStyle1);
}

function createHeaderRowSheet_1(sheet) {
    var headerRow = sheet.createRow(rowCounter_1);
    rowCounter_1++;
    
    sheet.setColumnWidth(0, 10000);
    sheet.setColumnWidth(1, 28000);
    
    var font = excelObject.createFont();
    font.setColor(new java.awt.Color(255 / 255.0, 255 / 255.0, 255 / 255.0, 1))
    font.setBold(true)
    
    cellStyle1 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(68, 114, 196), Constants.SOLID_FOREGROUND);
    //cellStyle2 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(155, 194, 230), Constants.SOLID_FOREGROUND);
    //cellStyle2.setLocked(true)
    //cellStyle3 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(255, 217, 102), Constants.SOLID_FOREGROUND);

    var newCell = headerRow.createCell(0);
    newCell.setCellValue("Operating Model");
    newCell.setCellStyle(cellStyle1);

    //var newCell = headerRow.createCell(1);
    //newCell.setCellValue("Link");
    //newCell.setCellStyle(cellStyle1);

}

function createHeaderRowSheet_2(sheet) {
    var headerRow = sheet.createRow(rowCounter_2);
    rowCounter_2++;
    
    sheet.setColumnWidth(0, 10000);
    sheet.setColumnWidth(1, 14000);
    sheet.setColumnWidth(2, 28000);
    
    //sheet.setDefaultColumnWidth(22)
    
    var font = excelObject.createFont();
    font.setColor(new java.awt.Color(255 / 255.0, 255 / 255.0, 255 / 255.0, 1))
    font.setBold(true)
    
    cellStyle1 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(68, 114, 196), Constants.SOLID_FOREGROUND);
    //cellStyle2 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(155, 194, 230), Constants.SOLID_FOREGROUND);
    //cellStyle2.setLocked(true)
    //cellStyle3 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(255, 217, 102), Constants.SOLID_FOREGROUND);

    var newCell = headerRow.createCell(0);
    newCell.setCellValue("Model Type");
    newCell.setCellStyle(cellStyle1);

    var newCell = headerRow.createCell(1);
    newCell.setCellValue("Model");
    newCell.setCellStyle(cellStyle1);
    
    //var newCell = headerRow.createCell(2);
    //newCell.setCellValue("Link");
    //newCell.setCellStyle(cellStyle1);

}

function createHeaderRowSheet_3(sheet) {
    var headerRow = sheet.createRow(rowCounter_3);
    rowCounter_3++;
    
    sheet.setColumnWidth(0, 10000);
    sheet.setColumnWidth(1, 10000);
    sheet.setColumnWidth(2, 10000);
    sheet.setColumnWidth(3, 10000);
    
    //sheet.setDefaultColumnWidth(22)
    
    var font = excelObject.createFont();
    font.setColor(new java.awt.Color(255 / 255.0, 255 / 255.0, 255 / 255.0, 1))
    font.setBold(true)
    
    cellStyle1 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(68, 114, 196), Constants.SOLID_FOREGROUND);
    //cellStyle2 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(155, 194, 230), Constants.SOLID_FOREGROUND);
    //cellStyle2.setLocked(true)
    //cellStyle3 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(255, 217, 102), Constants.SOLID_FOREGROUND);

    var newCell = headerRow.createCell(0);
    newCell.setCellValue("Process Owner");
    newCell.setCellStyle(cellStyle1);

    var newCell = headerRow.createCell(1);
    newCell.setCellValue("E2E Domain");
    newCell.setCellStyle(cellStyle1);
    
    var newCell = headerRow.createCell(2);
    newCell.setCellValue("Process Name");
    newCell.setCellStyle(cellStyle1);
    
    var newCell = headerRow.createCell(3);
    newCell.setCellValue("Model type");
    newCell.setCellStyle(cellStyle1);
}

function createHeaderRowSheet_4(sheet) {
    var headerRow = sheet.createRow(rowCounter_4);
    rowCounter_4++;
    
    sheet.setColumnWidth(0, 10000);
    sheet.setColumnWidth(1, 10000);
    sheet.setColumnWidth(2, 10000);
    
    //sheet.setDefaultColumnWidth(22)
    
    var font = excelObject.createFont();
    font.setColor(new java.awt.Color(255 / 255.0, 255 / 255.0, 255 / 255.0, 1))
    font.setBold(true)
    
    cellStyle1 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(68, 114, 196), Constants.SOLID_FOREGROUND);
    //cellStyle2 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(155, 194, 230), Constants.SOLID_FOREGROUND);
    //cellStyle2.setLocked(true)
    //cellStyle3 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(255, 217, 102), Constants.SOLID_FOREGROUND);

    var newCell = headerRow.createCell(0);
    newCell.setCellValue("Template");
    newCell.setCellStyle(cellStyle1);

    var newCell = headerRow.createCell(1);
    newCell.setCellValue("Domain");
    newCell.setCellStyle(cellStyle1);
    
    var newCell = headerRow.createCell(2);
    newCell.setCellValue("Main process");
    newCell.setCellStyle(cellStyle1);

}

function print_0(o_regionName, o_regionLink, o_countryName, o_countryLink, o_legalEntityNum, o_legalEntityName, o_sheet){
    var newRow = o_sheet.createRow(rowCounter_0);
    rowCounter_0++

    var newCell = newRow.createCell(0);
    //newCell.setCellValue(o_modelName);
    newCell.setHyperlink( o_regionName, o_regionLink )
    //newCell.setCellStyle(cellStyle);
    
    var newCell = newRow.createCell(1);
    //newCell.setCellValue(o_modelName);
    newCell.setHyperlink( o_countryName, o_countryLink )
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(2);
    newCell.setCellValue(o_legalEntityNum);
    //newCell.setCellStyle(cellStyle);
    
    var newCell = newRow.createCell(3);
    newCell.setCellValue(o_legalEntityName);
    //newCell.setCellStyle(cellStyle);
}

function print_1(o_modelName, o_link, o_sheet){
    var newRow = o_sheet.createRow(rowCounter_1);
    rowCounter_1++

    var newCell = newRow.createCell(0);
    //newCell.setCellValue(o_modelName);
    newCell.setHyperlink( o_modelName, o_link )
    //newCell.setCellStyle(cellStyle);

    //var newCell = newRow.createCell(1);
    //newCell.setCellValue(o_link);
    //newCell.setCellStyle(cellStyle);

}

function print_2(o_modelType, o_modelName, o_link, o_sheet){
    var newRow = o_sheet.createRow(rowCounter_2);
    rowCounter_2++
    
    var newCell = newRow.createCell(0);
    newCell.setCellValue(o_modelType);
    //newCell.setCellStyle(cellStyle);
    
    var newCell = newRow.createCell(1);
    //newCell.setCellValue(o_modelName);
    newCell.setHyperlink( o_modelName, o_link )
    //newCell.setCellStyle(cellStyle);

    //var newCell = newRow.createCell(2);
    //newCell.setCellValue(o_link);
    //newCell.setCellStyle(cellStyle);
}

function print_3(o_processOwner, o_e2eDomain, o_processName, o_processLink, o_modelType, o_sheet){
    var newRow = o_sheet.createRow(rowCounter_3);
    rowCounter_3++
    
    var font = excelObject.createFont();
    cellStyle = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT,Constants.C_TRANSPARENT, Constants.SOLID_FOREGROUND);

    
    var newCell = newRow.createCell(0);
    newCell.setCellValue(o_processOwner);
    //newCell.setCellStyle(cellStyle);
    
    var newCell = newRow.createCell(1);
    newCell.setCellValue(o_e2eDomain);
    //newCell.setCellStyle(cellStyle);
    
    var newCell = newRow.createCell(2);
    //newCell.setCellValue(o_modelName);
    newCell.setHyperlink( o_processName, o_processLink )
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(3);
    newCell.setCellValue(o_modelType);
    newCell.setCellStyle(cellStyle); //if removed cell is white ( BUG )
}

function print_4(o_templateName, o_e2eDomainAtt, o_modelName, o_modelLink, o_sheet){
    var newRow = o_sheet.createRow(rowCounter_4);
    rowCounter_4++
    
    //var font = excelObject.createFont();
    //cellStyle = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT,Constants.C_TRANSPARENT, Constants.SOLID_FOREGROUND);

    var newCell = newRow.createCell(0);
    newCell.setCellValue(o_templateName);
    //newCell.setCellStyle(cellStyle);
    
    var newCell = newRow.createCell(1);
    newCell.setCellValue(o_e2eDomainAtt);
    //newCell.setCellStyle(cellStyle);
    
    var newCell = newRow.createCell(2);
    //newCell.setCellValue(o_modelName);
    newCell.setHyperlink(o_modelName, o_modelLink )
    //newCell.setCellStyle(cellStyle);

}

function RGB(r, g, b) {
    return (new java.awt.Color(r / 255.0, g / 255.0, b / 255.0, 1)).getRGB() & 0xFFFFFF
}

function getPortalLink(model) {
    //http://<servername>/#<tenant>/item/c.<factSheet>.<DBName>.<GUID>.<Version>

    var oDatabase = ArisData.getActiveDatabase();
    var sTenantName = ArisData.getTenantName();
    var serverName = oDatabase.ServerName();
    var sDBName = oDatabase.Name(nLocale);
    sDBName = sDBName.replace(" ", "%20");
    var mGUID = model.GUID()

    var linkURL = "http://" + serverName + "/#" + sTenantName + "/item/c." + sDBName + "." + mGUID + ".-1"
    return linkURL
}

function getSheetByName(sheetName, excelObj){
    var allSheets = excelObj.getSheets()
    
    for(var i=0; i<allSheets.length; i++){
        if(allSheets[i].getName() == sheetName){
            return allSheets[i]
        }
    }
    return null; //sheet not found
}

function removeSheetContent(sheet){
    var rows = sheet.getRows()
    for(var i=0; i<rows.length; i++){
        sheet.removeRow(rows[i])
    }
}