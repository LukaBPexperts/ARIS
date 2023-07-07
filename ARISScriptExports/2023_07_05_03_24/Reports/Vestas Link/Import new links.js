var nLocale = Context.getSelectedLanguage()
var selectedDb = ArisData.getSelectedDatabases()

var outputFile = Context.createOutputObject()

var models = new Array()

var excelFile = getExcelFile() //get excell file
var sheets = excelFile.getSheets() //get all sheets

for (var i = 0; i < sheets.length; i++) {
    //check sheet name, check column name


    if (sheets[i].getName() == "Objects") {

        //get number of rows in excel table
        var rowCount = 0
        while (sheets[i].getCell(rowCount, 0) != null) {
            rowCount++
        }

        for (var j = 1; j < rowCount; j++) {
            Context.writeStatus("1. Reading excel file row " + (j + 1) + "/" + rowCount)
            var objGUID = oldLink_1 = sheets[i].getCell(j, 0).getCellValue()
            var s_objDef = selectedDb[0].FindGUID(objGUID)

            var value_ObjOldLink_1 = s_objDef.Attribute(Constants.AT_EXT_1, nLocale).getValue()
            var value_OldLink_1 = sheets[i].getCell(j, 3).getCellValue()
            var value_NewLink_1 = sheets[i].getCell(j, 8).getCellValue()

            if (objOldLink_1.includes(oldLink_1)) {
                //8df395b0-44af-11ed-2ba6-0a002700001a last change link 
                var updateLink1Att = s_objDef.Attribute(Constants.AT_EXT_1, nLocale).setValue(value_NewLink_1)
                var updateLink1DateAtt = s_objDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("8df395b0-44af-11ed-2ba6-0a002700001a"), nLocale).setValue(getTodayDate())
                
                outputFile.OutputTxt("UPDATED");
            }
            else{
                outputFile.OutputTxt("NOT UPDATED");
            }
        }
    }
    if (sheets[i].getName() == "Models") {
        //CODE HERE
    }
}

outputFile.WriteReport();

function getExcelFile() {
    var sdefname = ""
    sdefext = __toString("*.xls!!Excel|*.xls; *.xlsx||")
    var sdefdir = ""
    var stitle = "Chose excel file"

    var files = Dialogs.BrowseForFiles(sdefname, sdefext, sdefdir, stitle, 0)

    var excelFile = Context.getExcelReader(files[0].getData())

    return excelFile;
}

function getTodayDate() {

    var now = new Date();
    var formatter = new java.text.SimpleDateFormat('MM/dd/YYYY');
    var result = formatter.format(now);

    return result
}