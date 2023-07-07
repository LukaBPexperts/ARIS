//*****************************************************************************************************************************
//*******************************************SCRIPT PARAMETERS*****************************************************************
//*****************************************************************************************************************************
var groupPath = "SKF/0. Content Library/Applications" //group path where the new objects are going to be created, and temp model stored  /SKF/0. Content Library/Applications
var importAtt = true //true if object attributes should also be imported
var sheetName = "Application" //name of the sheet used for import

//APPLICATION FAMILY MODEL
var appFamilyModelGUID = "55ea60e0-85d3-11ec-1a81-0022487eead8"
var familyModelObjSymbol = Constants.ST_APPL_SYS_CLS
var familyModelType = ArisData.ActiveFilter().UserDefinedModelTypeNum("2c1d6381-d494-11e5-6fc6-005056a95aef")

var objDefTypeSource = Constants.OT_APPL_SYS_CLS //source objDef type
var objDefTypeTarget = Constants.OT_APPL_SYS_CLS //target objDef type
var cxnType = Constants.CT_CAN_SUBS_2 //connection type

var sourceIDColumn = 6 //source object ID column
var targetIDColumn = 2 // target object ID column

//TEMP MODEL
var tempModelType =  ArisData.ActiveFilter().UserDefinedModelTypeNum("2c1d6381-d494-11e5-6fc6-005056a95aef") //dummy model type, this model is used to create connection object
var tempSourceOccSymbol = Constants.ST_APPL_SYS_CLS //source object symbol used in the dummy model
var tempTargetOccSymbol = ArisData.ActiveFilter().UserDefinedSymbolTypeNum("12e88be1-e5cb-11e5-6fc6-005056a95aef") //target object symbol used in the dummy model

//Attributes maping
var idAttNum = ArisData.ActiveFilter().UserDefinedAttributeTypeNum("5becae21-d204-11e8-52ce-005056b341b8") // Correlation ID
var apmNumberAttNum = ArisData.ActiveFilter().UserDefinedAttributeTypeNum("36348e21-a6ec-11ed-6f81-0022487eead8")  //APM Number	

var capabilityL0AttNum = ArisData.ActiveFilter().UserDefinedAttributeTypeNum("85b71c21-b1f4-11ed-6f81-0022487eead8") //Capability L0	

var descriptionAttNum = Constants.AT_DESC 

var cmoTechLifeCycleStatusAttNum = ArisData.ActiveFilter().UserDefinedAttributeTypeNum("afdeb5a0-7e04-11ec-1a81-0022487eead8")
var funcSysOwnerAttNum = ArisData.ActiveFilter().UserDefinedAttributeTypeNum("f0fe5c90-c9ed-11e2-5475-005056c00008")
var itSysOwnerAttNum = ArisData.ActiveFilter().UserDefinedAttributeTypeNum("00374dc0-c9ee-11e2-5475-005056c00008")

//*****************************************************************************************************************************
//*****************************************************************************************************************************
//*****************************************************************************************************************************
var giStandardColumnWidth = 15000;
var giRowCounter = 0;
var giRowCounter2 = 0;
var lDataRowStyle = null;
var gCurrentFilter;

gExcelObject = Context.createExcelWorkbook(Context.getSelectedFile());

// set style header
var lHeaderFont = gExcelObject.createFont();
lHeaderFont.setBoldweight(Constants.XL_FONT_WEIGHT_BOLD);
var headLiBorderColor = Constants.C_GRAY;
var headerStyle = gExcelObject.createCellStyle(lHeaderFont, 1, 1, 1, 1, headLiBorderColor, headLiBorderColor, headLiBorderColor, headLiBorderColor, Constants.ALIGN_CENTER, Constants.VERTICAL_CENTER, Constants.C_TRANSPARENT, Constants.C_LIGHT_BLUE, Constants.SOLID_FOREGROUND);
// set style rows
var crFont = gExcelObject.createFont();
crFont.setBoldweight(Constants.XL_FONT_WEIGHT_NORMAL);
var crLiBorderColor = Constants.C_GRAY;
var crDataRowStyle = gExcelObject.createCellStyle(crFont, 1, 1, 1, 1, crLiBorderColor, crLiBorderColor, crLiBorderColor, crLiBorderColor, Constants.ALIGN_LEFT, Constants.VERTICAL_TOP, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.SOLID_FOREGROUND);
crDataRowStyle.setWrapText(true);

printRowHeadersSheet1(gExcelObject);
printRowHeadersSheet2(gExcelObject);

//************************************
var nLocale = Context.getSelectedLanguage()
var selectedDb = ArisData.getSelectedDatabases()

var group = selectedDb[0].Group(groupPath, nLocale)

var excelFile = getExcelFile() //get excell file
var sheets = excelFile.getSheets() //get all sheets


var sheet = getSheetByName(sheets, sheetName)

var rowCount = 0
if(sheet != null){
var rowCount =  sheet.getRows().length
}

for (var i = 1; i < rowCount; i++) {
    Context.writeStatus("1. Reading excel file row " + (i + 1) + "/" + rowCount)
    
    var idSource = ""
    var idTarget = ""
    var apmNumberTarget =""
    var sourceObjDef = null
    var targetObjDef = null
    
    var newSource = false
    var newTarget = false
    
    try{
    var idSource = sheet.getCell(i, sourceIDColumn).getCellValue().trim() //get ID (name) from excell
    var idTarget = sheet.getCell(i, targetIDColumn).getCellValue().trim() //get ID from excell
    
    var apmNumberColumnNum = getColumnNum(sheet, "Number")
    var apmNumberTarget = sheet.getCell(i, apmNumberColumnNum).getCellValue().trim() //get APM Number from excell
    
    }catch(e){}
    
    if(idTarget != "" || apmNumberTarget != ""){
       
       if(idSource != ""){
            var sourceObjDef = selectedDb[0].Find(Constants.SEARCH_OBJDEF, objDefTypeSource, Constants.AT_NAME, nLocale, idSource, Constants.SEARCH_CMP_EQUAL)[0] //search family  by the name       
       }       
       if(idTarget != ""){
            var targetObjDef = selectedDb[0].Find(Constants.SEARCH_OBJDEF, objDefTypeTarget, idAttNum, nLocale, idTarget, Constants.SEARCH_CMP_EQUAL)[0]
       }
       if(apmNumberTarget != "" && targetObjDef == null){
            var targetObjDef = selectedDb[0].Find(Constants.SEARCH_OBJDEF, objDefTypeTarget, apmNumberAttNum, nLocale, apmNumberTarget, Constants.SEARCH_CMP_EQUAL)[0]
       }
    
        if (idSource != "" && !checkObjDef(sourceObjDef)) {
            var nameSource = sheet.getCell(i, 6).getCellValue() //get ID from excell
            var sourceObjDef = group.CreateObjDef(objDefTypeSource, nameSource, Context.getSelectedLanguage())

            //create object in App family model
            /* 
            var appFamilyModel = selectedDb[0].FindGUID(appFamilyModelGUID)
            try{
                var familyModelObjOcc = appFamilyModel.createObjOcc(familyModelObjSymbol, sourceObjDef, 100, 100, true, true)
                var layoutParameters = ArisData.getModelOptions().getLayoutParameters(familyModelType)
    
                //layoutParameters.setArrangeSatellites(2)
                layoutParameters.setSpacingX(100)
                layoutParameters.setSpacingY(100)
                layoutParameters.setQuality(5)
    
                //currentModel.setAttrOccHandling( ArisData.getModelOptions().getTextAttributeHandling() )
                appFamilyModel.doLayout()
            }catch(ex){} 
            */

            //var updateSourceIDAtt = sourceObjDef.Attribute(idAttNum, nLocale).setValue(idSource)
            var newSource = true
            printRowSheet2(sourceObjDef.Name(nLocale), idSource, "New Object Created", "", "", gExcelObject)
        }   
       
        if (!checkObjDef(targetObjDef)) {
            var nameTarget = sheet.getCell(i, 1).getCellValue() //get ID from excell
            var targetObjDef = group.CreateObjDef(objDefTypeTarget, nameTarget, Context.getSelectedLanguage())
            var updateTargetIDAtt = targetObjDef.Attribute(idAttNum, nLocale).setValue(idTarget)
            var updateTargetApmNumberAtt = targetObjDef.Attribute(apmNumberAttNum, nLocale).setValue(apmNumberTarget)
            var newTarget = true
            printRowSheet2(targetObjDef.Name(nLocale), idTarget, "New Object Created", "", "", gExcelObject)
        }
        
        //check if container model exist
        var tempModel = selectedDb[0].Find(Constants.SEARCH_MODEL, tempModelType, Constants.AT_NAME, nLocale, "Application Container", Constants.SEARCH_CMP_EQUAL)[0] 
        
        if(tempModel == null){
            var tempModel = group.CreateModel(tempModelType, "Application Container", nLocale)
        }
        
        //check if connection already exist
        if(idSource != "" && !checkConnection(sourceObjDef, targetObjDef, tempModel)){
        
        //create connection
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
        //var delModel = group.Delete(tempModel)
        
        //apply layout
        //var layoutRes = tempModel.doLayout()
        }
        else if (checkConnection(sourceObjDef, targetObjDef, tempModel)){
        //connection already exist
        printRowSheet1(sourceObjDef.Name(nLocale), idSource, targetObjDef.Name(nLocale), idTarget, "Cxn already exist", gExcelObject)
        }
        else{
            var targetObjOcc = tempModel.createObjOcc(tempTargetOccSymbol, targetObjDef, 200, 200) //add to check if occurence already exist in model when creating connection
            printRowSheet1("", idSource, targetObjDef.Name(nLocale), idTarget, "Cxn is not created. Application family object is missing", gExcelObject)
        }
        
        //IMPORT ATTRIBUTES FOR SOURCE AND TARGET OBJECT
        if(importAtt == true){
            
            var capabilityL0AttValue = ""
            var apmNumberAttValue = ""
            var descriptionAttTarget = ""
            var cmoTechLifeCycleStatusAttTarget = ""
            var funcSysOwnerAttTarget = ""
            var itSysOwnerAttTarget = ""
            
            //var nameColumnNumTarget = getColumnNum(sheet, "Name")
            var descColumnNum = getColumnNum(sheet, "Description")  
            var cmoTechStatusColumnNum = getColumnNum(sheet, "Application Lifecycle stage")  
            var funcSysOwnerColumnNum = getColumnNum(sheet, "Business Owner")  
            var itSysOwnerColumnNum = getColumnNum(sheet, "IT Application Owner")  
            var apmNumberColumnNum = getColumnNum(sheet, "Number")
            
            var capabilityL0ColumnNum = getColumnNum(sheet, "Business process")
            
            try{            
            var descriptionAttTarget = sheet.getCell(i, descColumnNum).getCellValue() 
            var cmoTechLifeCycleStatusAttTarget = sheet.getCell(i, cmoTechStatusColumnNum).getCellValue()
            var funcSysOwnerAttTarget = sheet.getCell(i, funcSysOwnerColumnNum).getCellValue()
            var itSysOwnerAttTarget = sheet.getCell(i, itSysOwnerColumnNum).getCellValue()
            
            var apmNumberAttValue = sheet.getCell(i, apmNumberColumnNum).getCellValue()
            
            var capabilityL0AttValue = sheet.getCell(i, capabilityL0ColumnNum).getCellValue()
            
            }catch(e){}

            updateAttValue(targetObjDef, descriptionAttNum, descriptionAttTarget)
            updateAttValue(targetObjDef, cmoTechLifeCycleStatusAttNum, cmoTechLifeCycleStatusAttTarget)          
            updateAttValue(targetObjDef, funcSysOwnerAttNum, funcSysOwnerAttTarget)
            updateAttValue(targetObjDef, itSysOwnerAttNum, itSysOwnerAttTarget)
            
            updateAttValue(targetObjDef, apmNumberAttNum, apmNumberAttValue)
            updateAttValue(targetObjDef, capabilityL0AttNum, capabilityL0AttValue)
        }
    }
}


//add function which check if anything has been removed from APM 
//get all application objects from DB, and compare them with excell file

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
    //var lHeaderFont = pExcelWorkBook.createFont();
    //lHeaderFont.setBoldweight(Constants.XL_FONT_WEIGHT_BOLD);
    //var headLiBorderColor = Constants.C_GRAY;

    //headerStyle = pExcelWorkBook.createCellStyle(lHeaderFont, 1, 1, 1, 1, headLiBorderColor, headLiBorderColor, headLiBorderColor, headLiBorderColor, Constants.ALIGN_CENTER, Constants.VERTICAL_CENTER, Constants.C_TRANSPARENT, Constants.C_LIGHT_BLUE, Constants.SOLID_FOREGROUND);

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
    //var lHeaderFont = pExcelWorkBook.createFont();
    //lHeaderFont.setBoldweight(Constants.XL_FONT_WEIGHT_BOLD);
    //var headLiBorderColor = Constants.C_GRAY;

    //headerStyle = pExcelWorkBook.createCellStyle(lHeaderFont, 1, 1, 1, 1, headLiBorderColor, headLiBorderColor, headLiBorderColor, headLiBorderColor, Constants.ALIGN_CENTER, Constants.VERTICAL_CENTER, Constants.C_TRANSPARENT, Constants.C_LIGHT_BLUE, Constants.SOLID_FOREGROUND);

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
    
    var lCellL5 = lHeaderRow.createCell(4);
    lCellL5.setCellValue("New Value");
    lCellL5.setCellStyle(headerStyle);

    
    giRowCounter2++
}

function printRowSheet1(sourceObjName, sourceObjID, targetObjName, targetObjID, oStatus, pExcelWorkBook) {
    try {

        //var crFont = pExcelWorkBook.createFont();
        //crFont.setBoldweight(Constants.XL_FONT_WEIGHT_NORMAL);
        //var crLiBorderColor = Constants.C_GRAY;

        //crDataRowStyle = pExcelWorkBook.createCellStyle(crFont, 1, 1, 1, 1, crLiBorderColor, crLiBorderColor, crLiBorderColor, crLiBorderColor, Constants.ALIGN_LEFT, Constants.VERTICAL_TOP, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.SOLID_FOREGROUND);

        //crDataRowStyle.setWrapText(true);

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

function printRowSheet2(objectName, objectID, oStatus, oOldValue, oNewValue, pExcelWorkBook) {
    try {

        //var crFont = pExcelWorkBook.createFont();
        //crFont.setBoldweight(Constants.XL_FONT_WEIGHT_NORMAL);
        //var crLiBorderColor = Constants.C_GRAY;

        //crDataRowStyle = pExcelWorkBook.createCellStyle(crFont, 1, 1, 1, 1, crLiBorderColor, crLiBorderColor, crLiBorderColor, crLiBorderColor, Constants.ALIGN_LEFT, Constants.VERTICAL_TOP, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.SOLID_FOREGROUND);

        //crDataRowStyle.setWrapText(true);

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
        
        var lCell4 = lModelRow.createCell(4);
        lCell4.setCellValue(oNewValue);
        lCell4.setCellStyle(crDataRowStyle);


       
        giRowCounter2++;
    } catch (ex) {
        Dialogs.MsgBox("Error : " + ex);
    }
}

function getSheetByName(sheets, sheetName){
    var returnSheet = null
    for (var i = 0; i < sheets.length; i++){       
        if(sheets[i].getName() == sheetName){    
            returnSheet = sheets[i]  
        }       
    }
    return returnSheet
}

function getColumnNum(sheet, columnName){
var temp = null;
for(var i = 0; i < 50; i++){
    if(sheet.getCell(0, i).getCellValue().trim().toLowerCase() == columnName.trim().toLowerCase()){
        var temp = i;
        break;
    }
}
return temp;
}

function updateAttValue(objDef, attNum, attValue){
if(objDef.Attribute(attNum, nLocale).getValue() != attValue){
    var oldValue = objDef.Attribute(attNum, nLocale).getValue()
    var updateAtt = objDef.Attribute(attNum, nLocale).setValue(attValue) 
    if(updateAtt){
        printRowSheet2(objDef.Name(nLocale), idSource, objDef.Attribute(attNum, nLocale).Type() + " - Attribute Updated", oldValue, attValue, gExcelObject)
    }
    else{
        printRowSheet2(objDef.Name(nLocale), idSource, "ERROR: " + objDef.Attribute(attNum, nLocale).Type() + " - Attribute Not Updated", oldValue, attValue, gExcelObject)
    }
}
}

function checkConnection(sourceObjDef, targetObjDef, modelFilter){
if(sourceObjDef != null){
var cxnList = sourceObjDef.CxnList([modelFilter])

for(var i=0; i<cxnList.length; i++){
    if(cxnList[i].TargetObjDef().GUID() == targetObjDef.GUID()){
        return true
    }
}
}
return false
}