//*****************************************************************************************************************************
//*******************************************SCRIPT PARAMETERS*****************************************************************
//*****************************************************************************************************************************
var groupPath = "SKF/4. Organization" //group path where the new objects are going to be created, and temp model stored
var importAtt = true //true if object attributes should also be imported

var sourceIDColumn = 4 //source object ID column
var targetIDColumn = 1 // target object ID column

var objDefTypeSource = Constants.OT_ORG_UNIT //source objDef type
var objDefTypeTarget = Constants.OT_ORG_UNIT //target objDef type
var cxnType = Constants.CT_IS_CRT_BY //connection type

//TEMP MODEL
var tempModelType = Constants.MT_ORG_CHRT //dummy model type, this model is used to create connection object
var tempSourceOccSymbol = Constants.ST_ORG_UNIT_1 //source object symbol used in the dummy model
var tempTargetOccSymbol = Constants.ST_ORG_UNIT_1 //target object symbol used in the dummy model

//Attributes maping
var idAttNum = ArisData.ActiveFilter().UserDefinedAttributeTypeNum("e7f29411-56fc-11e5-15b5-b4b52f5aab7a")
var orgUnitTypeAttNum = ArisData.ActiveFilter().UserDefinedAttributeTypeNum("31f9d701-5d1d-11e5-15b5-b4b52f5aab7a")
var countryAttNum = ArisData.ActiveFilter().UserDefinedAttributeTypeNum("73c14d31-6754-11e5-15b5-b4b52f5aab7a")
var organizationalIDAttNum = ArisData.ActiveFilter().UserDefinedAttributeTypeNum("e7f29411-56fc-11e5-15b5-b4b52f5aab7a")
var descriptionAttNum = Constants.AT_DESC 

//*****************************************************************************************************************************
//*****************************************************************************************************************************
//*****************************************************************************************************************************
var giStandardColumnWidth = 15000;
var giRowCounter = 0;
var giRowCounter2 = 0;
var lDataRowStyle = null;
var gCurrentFilter;


gExcelObject = Context.createExcelWorkbook(Context.getSelectedFile());

printRowHeadersSheet1(gExcelObject);
printRowHeadersSheet2(gExcelObject);

//************************************
var nLocale = Context.getSelectedLanguage()
var selectedDb = ArisData.getSelectedDatabases()

var group = selectedDb[0].Group(groupPath, nLocale)

var excelFile = getExcelFile() //get excell file
var sheets = excelFile.getSheets() //get all sheets

var rowCount =  sheets[0].getRows().length

for (var i = 1; i < rowCount; i++) {
    Context.writeStatus("1. Reading excel file row " + (i + 1) + "/" + rowCount)
    
    var idSource = ""
    var idTarget = ""
    var sourceObjDef = null
    var targetObjDef = null
    
    try{
    var idSource = sheets[0].getCell(i, sourceIDColumn).getCellValue().trim() //get ID from excell
    var idTarget = sheets[0].getCell(i, targetIDColumn).getCellValue().trim() //get ID from excell
    }catch(e){}
    
    if(idSource != "" && idTarget != ""){
        var sourceObjDef = selectedDb[0].Find(Constants.SEARCH_OBJDEF, objDefTypeSource, idAttNum, nLocale, idSource, Constants.SEARCH_CMP_EQUAL)[0]
        var targetObjDef = selectedDb[0].Find(Constants.SEARCH_OBJDEF, objDefTypeTarget, idAttNum, nLocale, idTarget, Constants.SEARCH_CMP_EQUAL)[0]
    
        if (!checkObjDef(sourceObjDef)) {
            var nameSource = sheets[0].getCell(i, 5).getCellValue() //get ID from excell
            var sourceObjDef = group.CreateObjDef(objDefTypeSource, nameSource, Context.getSelectedLanguage())
            var updateSourceIDAtt = sourceObjDef.Attribute(idAttNum, nLocale).setValue(idSource)
            printRowSheet2(sourceObjDef.Name(nLocale), idSource, "New Object Created", "", gExcelObject)
        }   
       
        if (!checkObjDef(targetObjDef)) {
            var nameTarget = sheets[0].getCell(i, 3).getCellValue() //get ID from excell
            var targetObjDef = group.CreateObjDef(objDefTypeTarget, nameTarget, Context.getSelectedLanguage())
            var updateTargetIDAtt = targetObjDef.Attribute(idAttNum, nLocale).setValue(idTarget)
            printRowSheet2(targetObjDef.Name(nLocale), idTarget, "New Object Created", "", gExcelObject)
        }
        
        //create connection
        var tempModel = group.CreateModel(tempModelType, "tempModel", nLocale)
        
        try{
        var sourceObjOcc = tempModel.createObjOcc(tempSourceOccSymbol, sourceObjDef, 100, 100)
        var targetObjOcc = tempModel.createObjOcc(tempTargetOccSymbol, targetObjDef, 200, 200)
        }catch(ex){}
               
        // var g_ntranspoint = new Array()
        // g_ntranspoint[0] = new PublicInterface.Point(sourceObjOcc.X(), sourceObjOcc.Y());
        // g_ntranspoint[1] = new PublicInterface.Point(targetObjOcc.X(), targetObjOcc.Y());
        // var temoCxnOcc = tempModel.CreateCxnOcc(sourceObjOcc, targetObjOcc, cxnType, g_ntranspoint)
        
        var pointList = new Array()
        var oPointSource = new java.awt.Point(sourceObjOcc.X(), sourceObjOcc.Y());
        var oPointTarget = new java.awt.Point(targetObjOcc.X(), targetObjOcc.Y());
        pointList.push(oPointSource)
        pointList.push(oPointTarget)       
        var tempCxnOcc = tempModel.CreateCxnOcc(sourceObjOcc, targetObjOcc, cxnType, pointList)
        
        if(tempCxnOcc.IsValid()){
            printRowSheet1(sourceObjDef.Name(nLocale), idSource, targetObjDef.Name(nLocale), idTarget, "Cxn created", gExcelObject)       
        }
        else{
            printRowSheet1(sourceObjDef.Name(nLocale), idSource, targetObjDef.Name(nLocale), idTarget, "Cxn is not created", gExcelObject)
        }
    
        //delete model after the object and cxn is created
        var delModel = group.Delete(tempModel)
        
        //IMPORT ATTRIBUTES FOR SOURCE AND TARGET OBJECT
        if(importAtt == true){
            var nameSource = ""
            var nameTarget = ""
            
            var locationAttSource = ""
            var countryAttTarget = ""
            var descriptionAttSource = ""
            var descriptionAttTarget = ""
            
            try{
            var nameSource = sheets[0].getCell(i, 5).getCellValue() //get Name from excell
            var nameTarget = sheets[0].getCell(i, 3).getCellValue() //get Name from excell
            
            var countryAttTarget = sheets[0].getCell(i, 8).getCellValue() 
            //var countryAttTarget = sheets[0].getCell(i, 5).getCellValue() 
            
            var descriptionAttTarget = sheets[0].getCell(i, 12).getCellValue() 
            //var descriptionAttTarget = sheets[0].getCell(i, 5).getCellValue() 
            }catch(e){}
            
            if(sourceObjDef.Attribute(Constants.AT_NAME, nLocale).getValue() != nameSource){
                var oldValue = sourceObjDef.Attribute(Constants.AT_NAME, nLocale).getValue()
                var updateNameAttSource = sourceObjDef.Attribute(Constants.AT_NAME, nLocale).setValue(nameSource)
                printRowSheet2(sourceObjDef.Name(nLocale), idSource, "Name Attribute Updated", oldValue, gExcelObject)
            }
            
            if(targetObjDef.Attribute(Constants.AT_NAME, nLocale).getValue() != nameTarget){
                var oldValue = targetObjDef.Attribute(Constants.AT_NAME, nLocale).getValue()
                var updateNameAttTarget = targetObjDef.Attribute(Constants.AT_NAME, nLocale).setValue(nameTarget)
                printRowSheet2(targetObjDef.Name(nLocale), idTarget, "Name Attribute Updated", oldValue, gExcelObject)
            }
            
            if(targetObjDef.Attribute(countryAttNum, nLocale).getValue() != countryAttTarget){
                var oldValue = targetObjDef.Attribute(countryAttNum, nLocale).getValue()
                var updateLocationAttTarget = targetObjDef.Attribute(countryAttNum, nLocale).setValue(countryAttTarget)
                printRowSheet2(targetObjDef.Name(nLocale), idTarget, "Country Attribute Updated", oldValue, gExcelObject)
            }
            
            if(targetObjDef.Attribute(descriptionAttNum, nLocale).getValue() != descriptionAttTarget){
                var oldValue = targetObjDef.Attribute(descriptionAttNum, nLocale).getValue()
                var updateDescTarget = targetObjDef.Attribute(descriptionAttNum, nLocale).setValue(descriptionAttTarget)
                printRowSheet2(targetObjDef.Name(nLocale), idTarget, "Description Attribute Updated", oldValue, gExcelObject)
            }
            
            if(targetObjDef.Attribute(orgUnitTypeAttNum, nLocale).getValue() != "Operating Unit" ){
                var oldValue = targetObjDef.Attribute(orgUnitTypeAttNum, nLocale).getValue()
                var updateOrgUnitTypeTarget = targetObjDef.Attribute(orgUnitTypeAttNum, nLocale).setValue(ArisData.ActiveFilter().UserDefinedAttributeValueTypeNum("31f9d701-5d1d-11e5-15b5-b4b52f5aab7a", "320f33cd-5d1d-11e5-15b5-b4b52f5aab7a"))
                printRowSheet2(targetObjDef.Name(nLocale), idTarget, "Attribute Organizational Unit Type Updated", oldValue, gExcelObject)
            }
            
        }
    }
}

gExcelObject.write();

function getExcelFile() {
    var sdefname = ""
    sdefext = __toString("*.xls!!Excel|*.xls; *.xlsx||")
    var sdefdir = ""
    var stitle = "Chose excel file"

    var files = Dialogs.BrowseForFiles(sdefname, sdefext, sdefdir, stitle, 0)

    var excelFile = Context.getExcelReader(files[0].getData())

    return excelFile;
}

function checkObjDef(objDef){
var validFlag = true

try{
    if (!objDef.IsValid()) {
        validFlag = false
    }

}catch(ex){}

if(objDef == null){
    validFlag = false
}

return validFlag
}

function printRowHeadersSheet1(pExcelWorkBook) {
    //var sDate = java.text.SimpleDateFormat("dd-MM-yyyy HH.mm").format(new java.util.Date());
    var sSheetName = "Cxns Log"
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
    lCellL1.setCellValue("Source Object Name");
    lCellL1.setCellStyle(headerStyle);

    //var lCellStyle = lCellL1.getCellStyle();
    //lCellStyle.setFont(lHeaderFont);

    var lCellL2 = lHeaderRow.createCell(1);
    lCellL2.setCellValue("Source Object ID");
    lCellL2.setCellStyle(headerStyle);

    var lCellL3 = lHeaderRow.createCell(2);
    lCellL3.setCellValue("Target Object Name");
    lCellL3.setCellStyle(headerStyle);

    var lCellL4 = lHeaderRow.createCell(3);
    lCellL4.setCellValue("Target Object ID");
    lCellL4.setCellStyle(headerStyle);
    
    var lCellL5 = lHeaderRow.createCell(4);
    lCellL5.setCellValue("Status");
    lCellL5.setCellStyle(headerStyle);

    
    giRowCounter++
}

function printRowHeadersSheet2(pExcelWorkBook) {
    //var sDate = java.text.SimpleDateFormat("dd-MM-yyyy HH.mm").format(new java.util.Date());
    var sSheetName = "Objects Log"
    var lMainSheet = pExcelWorkBook.createSheet(sSheetName);
    var lHeaderRow = lMainSheet.createRow(giRowCounter2);
    

    lMainSheet.setDefaultColumnWidth(30);

    /* Freeze top row alone */
    lMainSheet.createFreezePane(0, 1);

    // set style
    var lHeaderFont = pExcelWorkBook.createFont();
    lHeaderFont.setBoldweight(Constants.XL_FONT_WEIGHT_BOLD);
    var headLiBorderColor = Constants.C_GRAY;

    headerStyle = pExcelWorkBook.createCellStyle(lHeaderFont, 1, 1, 1, 1, headLiBorderColor, headLiBorderColor, headLiBorderColor, headLiBorderColor, Constants.ALIGN_CENTER, Constants.VERTICAL_CENTER, Constants.C_TRANSPARENT, Constants.C_LIGHT_BLUE, Constants.SOLID_FOREGROUND);

    var lCellL1 = lHeaderRow.createCell(0);
    lCellL1.setCellValue("Object Name");
    lCellL1.setCellStyle(headerStyle);

    //var lCellStyle = lCellL1.getCellStyle();
    //lCellStyle.setFont(lHeaderFont);

    var lCellL2 = lHeaderRow.createCell(1);
    lCellL2.setCellValue("Object ID");
    lCellL2.setCellStyle(headerStyle);

    var lCellL3 = lHeaderRow.createCell(2);
    lCellL3.setCellValue("Status");
    lCellL3.setCellStyle(headerStyle);
    
    var lCellL4 = lHeaderRow.createCell(3);
    lCellL4.setCellValue("Old Value");
    lCellL4.setCellStyle(headerStyle);

    
    giRowCounter2++
}

function printRowSheet1(sourceObjName, sourceObjID, targetObjName, targetObjID, oStatus, pExcelWorkBook) {
    try {

        var crFont = pExcelWorkBook.createFont();
        crFont.setBoldweight(Constants.XL_FONT_WEIGHT_NORMAL);
        var crLiBorderColor = Constants.C_GRAY;

        crDataRowStyle = pExcelWorkBook.createCellStyle(crFont, 1, 1, 1, 1, crLiBorderColor, crLiBorderColor, crLiBorderColor, crLiBorderColor, Constants.ALIGN_LEFT, Constants.VERTICAL_TOP, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.SOLID_FOREGROUND);

        crDataRowStyle.setWrapText(true);

        var lMainSheet = pExcelWorkBook.getSheetAt(0);

        // second row data
        var lModelRow = lMainSheet.createRow(giRowCounter);

        var lCell0 = lModelRow.createCell(0);
        lCell0.setCellValue(sourceObjName);
        lCell0.setCellStyle(crDataRowStyle);

        var lCell1 = lModelRow.createCell(1);
        lCell1.setCellValue(sourceObjID);
        lCell1.setCellStyle(crDataRowStyle);

        var lCell2 = lModelRow.createCell(2);
        lCell2.setCellValue(targetObjName);
        lCell2.setCellStyle(crDataRowStyle);

        // Input Name
        var lCellBPO = lModelRow.createCell(3);
        lCellBPO.setCellValue(targetObjID);
        lCellBPO.setCellStyle(crDataRowStyle);
        
        var lCellBPO = lModelRow.createCell(4);
        lCellBPO.setCellValue(oStatus);
        lCellBPO.setCellStyle(crDataRowStyle);


        // up counter
        giRowCounter++;
    } catch (ex) {
        Dialogs.MsgBox("Error : " + ex);
    }
}

function printRowSheet2(objectName, objectID, oStatus, oOldValue, pExcelWorkBook) {
    try {

        var crFont = pExcelWorkBook.createFont();
        crFont.setBoldweight(Constants.XL_FONT_WEIGHT_NORMAL);
        var crLiBorderColor = Constants.C_GRAY;

        crDataRowStyle = pExcelWorkBook.createCellStyle(crFont, 1, 1, 1, 1, crLiBorderColor, crLiBorderColor, crLiBorderColor, crLiBorderColor, Constants.ALIGN_LEFT, Constants.VERTICAL_TOP, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.SOLID_FOREGROUND);

        crDataRowStyle.setWrapText(true);

        var lMainSheet = pExcelWorkBook.getSheetAt(1);

        // second row data
        var lModelRow = lMainSheet.createRow(giRowCounter2);

        var lCell0 = lModelRow.createCell(0);
        lCell0.setCellValue(objectName);
        lCell0.setCellStyle(crDataRowStyle);

        var lCell1 = lModelRow.createCell(1);
        lCell1.setCellValue(objectID);
        lCell1.setCellStyle(crDataRowStyle);

        var lCell2 = lModelRow.createCell(2);
        lCell2.setCellValue(oStatus);
        lCell2.setCellStyle(crDataRowStyle);
        
        var lCell3 = lModelRow.createCell(3);
        lCell3.setCellValue(oOldValue);
        lCell3.setCellStyle(crDataRowStyle);


       
        giRowCounter2++;
    } catch (ex) {
        Dialogs.MsgBox("Error : " + ex);
    }
}