var lastChangeLink1 = "da4b8da1-44b4-11ed-0c1c-00505681a475"
var lastChangeLink2 = "ab54c300-44bd-11ed-0c1c-00505681a475"
var lastChangeLink3 = "b1efbda1-44bd-11ed-0c1c-00505681a475"
var lastChangeLink4 = "b95e0dd1-44bd-11ed-0c1c-00505681a475"
var lastChangeLink5 = "bf4a2ad1-44bd-11ed-0c1c-00505681a475"
var lastChangeLink6 = "c44f6a41-44bd-11ed-0c1c-00505681a475"

var nLocale = Context.getSelectedLanguage()
var selectedDb = ArisData.getSelectedDatabases()

var outputFile = Context.createOutputObject()

var models = new Array()

var excelFile = getExcelFile() //get excell file
var sheets = excelFile.getSheets() //get all sheets

for (var i = 0; i < sheets.length; i++) {
    //check sheet name, check column name


    if (sheets[i].getName() == "Output File 20220922_Design Obj") {

        //get number of rows in excel table
        var rowCount = 0
        while (sheets[i].getCell(rowCount, 0) != null) {
            rowCount++
        }

        for (var j = 1; j < rowCount; j++) {
            Context.writeStatus("1. Reading excel file row " + (j + 1) + "/" + rowCount)
            var objGUID = oldLink_1 = sheets[i].getCell(j, 0).getCellValue()
            var s_objDef = selectedDb[0].FindGUID(objGUID)

            try {
                if (sheets[i].getCell(0, 8).getCellValue() == "Updated Link1") { //CHANGE TO Link 1
                    var value_NewLink_1 = sheets[i].getCell(j, 8).getCellValue()
                }
            }
            catch (e) {
                var value_NewLink_1 = ""
            }
            try {
                if (sheets[i].getCell(0, 14).getCellValue() == "Updated Link 2") {
                    var value_NewLink_2 = sheets[i].getCell(j, 14).getCellValue()
                }
            }
            catch (e) {
                var value_NewLink_2 = ""
            }
            try {
                if (sheets[i].getCell(0, 20).getCellValue() == "Updated Link 3") {
                    var value_NewLink_3 = sheets[i].getCell(j, 20).getCellValue()
                }
            }
            catch (e) {
                var value_NewLink_3 = ""
            }
            try {
                if (sheets[i].getCell(0, 26).getCellValue() == "Updated Link 4") {
                    var value_NewLink_4 = sheets[i].getCell(j, 26).getCellValue()
                }
            }
            catch (e) {
                var value_NewLink_4 = ""
            }
            try {
                if (sheets[i].getCell(0, 32).getCellValue() == "Updated Link 5") {
                    var value_NewLink_5 = sheets[i].getCell(j, 32).getCellValue()
                }
            }
            catch (e) {
                var value_NewLink_5 = ""
            }
            try {
                if (sheets[i].getCell(0, 38).getCellValue() == "Updated Link 6") {
                    var value_NewLink_6 = sheets[i].getCell(j, 38).getCellValue()
                }
            }
            catch (e) {
                var value_NewLink_6 = ""
            }
            var updateLink1Att = false
            var updateLink2Att = false
            var updateLink3Att = false
            var updateLink4Att = false
            var updateLink5Att = false
            var updateLink6Att = false

            if (value_NewLink_1 != "" && value_NewLink_1 != null) {
                var updateLink1Att = s_objDef.Attribute(Constants.AT_EXT_1, nLocale).setValue(value_NewLink_1)
                var updateLink1DateAtt = s_objDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum(lastChangeLink1), nLocale).setValue(getTodayDate())
            }
            if (value_NewLink_2 != "" && value_NewLink_2 != null) {
                var updateLink2Att = s_objDef.Attribute(Constants.AT_EXT_2, nLocale).setValue(value_NewLink_1)
                var updateLink2DateAtt = s_objDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum(lastChangeLink2), nLocale).setValue(getTodayDate())
            }
            if (value_NewLink_3 != "" && value_NewLink_3 != null) {
                var updateLink3Att = s_objDef.Attribute(Constants.AT_EXT_3, nLocale).setValue(value_NewLink_1)
                var updateLink3DateAtt = s_objDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum(lastChangeLink3), nLocale).setValue(getTodayDate())
            }
            if (value_NewLink_4 != "" && value_NewLink_4 != null) {
                var updateLink4Att = s_objDef.Attribute(Constants.AT_LINK, nLocale).setValue(value_NewLink_1)
                var updateLink4DateAtt = s_objDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum(lastChangeLink4), nLocale).setValue(getTodayDate())
            }
            if (value_NewLink_5 != "" && value_NewLink_5 != null) {
                var updateLink5Att = s_objDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("fcba0a10-f523-11e6-004f-eabd1a8bb6a2"), nLocale).setValue(value_NewLink_1)
                var updateLink5DateAtt = s_objDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum(lastChangeLink5), nLocale).setValue(getTodayDate())
            }
            if (value_NewLink_6 != "" && value_NewLink_6 != null) {
                var updateLink6Att = s_objDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("6cb95000-19d2-11e7-004f-eabd1a8bb6a2"), nLocale).setValue(value_NewLink_1)
                var updateLink6DateAtt = s_objDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum(lastChangeLink6), nLocale).setValue(getTodayDate())
            }

            if (updateLink1Att == true) {
                outputFile.OutputTxt("OBJECT: " + objGUID + "/" + s_objDef.Name(nLocale) + "/" + "  Link 1 --> UPDATED" + "\n");
            }
            else {
                outputFile.OutputTxt("OBJECT: " + objGUID + "/" + s_objDef.Name(nLocale) + "/" + "  Link 1 --> NOT UPDATED" + "\n");
            }
            if (updateLink2Att == true) {
                outputFile.OutputTxt("OBJECT: " + objGUID + "/" + s_objDef.Name(nLocale) + "/" + "  Link 2 --> UPDATED" + "\n");
            }
            else {
                outputFile.OutputTxt("OBJECT: " + objGUID + "/" + s_objDef.Name(nLocale) + "/" + "  Link 2 --> NOT UPDATED" + "\n");
            }
            if (updateLink3Att == true) {
                outputFile.OutputTxt("OBJECT: " + objGUID + "/" + s_objDef.Name(nLocale) + "/" + "  Link 3 --> UPDATED" + "\n");
            }
            else {
                outputFile.OutputTxt("OBJECT: " + objGUID + "/" + s_objDef.Name(nLocale) + "/" + "  Link 3 --> NOT UPDATED" + "\n");
            }
            if (updateLink4Att == true) {
                outputFile.OutputTxt("OBJECT: " + objGUID + "/" + s_objDef.Name(nLocale) + "/" + "  Link 4 --> UPDATED" + "\n");
            }
            else {
                outputFile.OutputTxt("OBJECT: " + objGUID + "/" + s_objDef.Name(nLocale) + "/" + "  Link 4 --> NOT UPDATED" + "\n");
            }
            if (updateLink5Att == true) {
                outputFile.OutputTxt("OBJECT: " + objGUID + "/" + s_objDef.Name(nLocale) + "/" + "  Link 5 --> UPDATED" + "\n");
            }
            else {
                outputFile.OutputTxt("OBJECT: " + objGUID + "/" + s_objDef.Name(nLocale) + "/" + "  Link 5 --> NOT UPDATED" + "\n");
            }
            if (updateLink6Att == true) {
                outputFile.OutputTxt("OBJECT: " + objGUID + "/" + s_objDef.Name(nLocale) + "/" + "  Link 6 --> UPDATED" + "\n");
            }
            else {
                outputFile.OutputTxt("OBJECT: " + objGUID + "/" + s_objDef.Name(nLocale) + "/" + "  Link 6 --> NOT UPDATED" + "\n");
            }

            outputFile.OutputTxt("\n");
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