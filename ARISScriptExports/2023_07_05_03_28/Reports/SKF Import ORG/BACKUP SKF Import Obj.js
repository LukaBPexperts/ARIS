//*****************************************************************************************************************************
//*******************************************SCRIPT PARAMETERS*****************************************************************
//*****************************************************************************************************************************
var groupPath = "SKF/4. Organization/TestOrg" //group path where the new objects are going to be created, and temp model stored

var sourceIDColumn = 0 //source object ID column
var targetIDColumn = 2 // target object ID column

var objDefTypeSource = Constants.OT_ORG_UNIT //source objDef type
var objDefTypeTarget = Constants.OT_ORG_UNIT //target objDef type
var cxnType = Constants.CT_IS_CRT_BY //connection type

//TEMP MODEL
var tempModelType = Constants.MT_ORG_CHRT //dummy model type, this model is used to create connection object
var tempSourceOccSymbol = Constants.ST_ORG_UNIT_1 //source object symbol used in the dummy model
var tempTargetOccSymbol = Constants.ST_ORG_UNIT_1 //target object symbol used in the dummy model

//*****************************************************************************************************************************
//*****************************************************************************************************************************
//*****************************************************************************************************************************

var nLocale = Context.getSelectedLanguage()
var selectedDb = ArisData.getSelectedDatabases()

var group = selectedDb[0].Group(groupPath, nLocale)

var excelFile = getExcelFile() //get excell file
var sheets = excelFile.getSheets() //get all sheets

//get number of rows in excel table
// var rowCount = 0
// while (sheets[0].getCell(rowCount, 0) != null) {
    // rowCount++
// }

var rowCount =  sheets[0].getRows().length

for (var i = 1; i < rowCount; i++) {
    Context.writeStatus("1. Reading excel file row " + (i + 1) + "/" + rowCount)
    
    var idSource = ""
    var idTarget = ""
    
    try{
    var idSource = sheets[0].getCell(i, sourceIDColumn).getCellValue() //get ID from excell
    var idTarget = sheets[0].getCell(i, targetIDColumn).getCellValue() //get ID from excell
    }catch(e){}
    
    //check if object exist in DB, if not create it
    var sourceObjDef = selectedDb[0].FindGUID(idSource)
    var targetObjDef = selectedDb[0].FindGUID(idTarget)

    if (!sourceObjDef.IsValid()) {
        var nameSource = sheets[0].getCell(i, 1).getCellValue() //get ID from excell
        var sourceObjDef = group.CreateObjDef(objDefTypeSource, nameSource, Context.getSelectedLanguage())
    }

    if (!targetObjDef.IsValid()) {
        var nameTarget = sheets[0].getCell(i, 3).getCellValue() //get ID from excell
        var targetObjDef = group.CreateObjDef(objDefTypeTarget, nameTarget, Context.getSelectedLanguage())
    }

    //create connection
    var tempModel = group.CreateModel(tempModelType, "tempModel", nLocale)

    var sourceObjOcc = tempModel.createObjOcc(tempSourceOccSymbol, sourceObjDef, 100, 100)
    var targetObjOcc = tempModel.createObjOcc(tempTargetOccSymbol, targetObjDef, 200, 200)

    var g_ntranspoint = new Array()
    g_ntranspoint[0] = new PublicInterface.Point(sourceObjOcc.X(), sourceObjOcc.Y());
    g_ntranspoint[1] = new PublicInterface.Point(targetObjOcc.X(), targetObjOcc.Y());
    var temoCxnOcc = tempModel.CreateCxnOcc(sourceObjOcc, targetObjOcc, cxnType, g_ntranspoint)

    //delete model after the object and cxn is created
    var delModel = group.Delete(tempModel)
    
    //IMPORT ATTRIBUTES FOR SOURCE OBJECT
    var nameSource = ""
    var nameTarget = ""
    
    try{
    var nameSource = sheets[0].getCell(i, 1).getCellValue() //get Name from excell
    var nameTarget = sheets[0].getCell(i, 3).getCellValue() //get Name from excell
    }catch(e){}
    
    
    var updateNameAttSource = sourceObjDef.Attribute(Constants.AT_NAME, nLocale).setValue(nameSource)
    var updateNameAttTarget = targetObjDef.Attribute(Constants.AT_NAME, nLocale).setValue(nameTarget)
    
    //sourceObjDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("adde7a20-75f9-11e0-61d8-00237d347a8a"), nLocale).setValue("value")
}

function getExcelFile() {
    var sdefname = ""
    sdefext = __toString("*.xls!!Excel|*.xls; *.xlsx||")
    var sdefdir = ""
    var stitle = "Chose excel file"

    var files = Dialogs.BrowseForFiles(sdefname, sdefext, sdefdir, stitle, 0)

    var excelFile = Context.getExcelReader(files[0].getData())

    return excelFile;
}