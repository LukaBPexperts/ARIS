var nLocale = Context.getSelectedLanguage()
var selectedDb = ArisData.getSelectedDatabases()

var cxnsForDeletion = new Array()

var excelFile = getExcelFile() //get excell file
var sheets = excelFile.getSheets() //get all sheets

for (var i = 0; i < sheets.length; i++) {
    //get number of rows in excel table
    var rowCount = 0
    while (sheets[i].getCell(rowCount, 0) != null) {
        rowCount++
    }

    //get marked cxns from file and check #occ
    for (var j = 1; j < rowCount; j++) {
        Context.writeStatus("1. Reading excel file row " + (j + 1) + "/" + rowCount)
        if (sheets[i].getCell(j, 12).getCellValue() == "YES" && sheets[i].getCell(j, 8).getCellValue() == 0) {
            cxnsForDeletion.push(sheets[i].getCell(j, 4).getCellValue())
        }
    }
}

cxnsForDeletion = ArisData.Unique(cxnsForDeletion)
//check if all cxns exist in db
var check = checkCxns(cxnsForDeletion)

if (check == true) {
    var btnPressed = Dialogs.MsgBox(cxnsForDeletion.length + " connections will be deleted. Do you want to continue?", Constants.MSGBOX_ICON_QUESTION | Constants.MSGBOX_BTN_YESNO + 512, "Message from Report")
    if (btnPressed == Constants.MSGBOX_RESULT_YES) {
        //delete connections
        for (var i = 0; i < cxnsForDeletion.length; i++) {
            Context.writeStatus("3. Deleting connection " + (i + 1) + "/" + (cxnsForDeletion.length))
            var cxnObject = selectedDb[0].FindGUID(cxnsForDeletion[i])
            
            var deletedCxn = cxnObject.Delete(true)
            
        }
    }
}

function checkCxns(cxnsForDeletion) {

    for (var i = 0; i < cxnsForDeletion.length; i++) {
        Context.writeStatus("2. Analyzing connection " + (i + 1) + "/" + (cxnsForDeletion.length))
        var checkFlag = selectedDb[0].FindGUID(cxnsForDeletion[i])

        if (checkFlag.IsValid() == false) {
            Dialogs.MsgBox(cxnsForDeletion[i] + " Connection GUID is not found in the database. No connections will be deleted", Constants.MSGBOX_ICON_ERROR | Constants.MSGBOX_BTN_OK + 512, "Excel File Error")
            return false;

        }
    }

    return true;
}

function getExcelFile() {
    var sdefname = ""
    sdefext = __toString("*.xls!!Excel|*.xls; *.xlsx||")
    var sdefdir = ""
    var stitle = "Chose Template excel file"

    var files = Dialogs.BrowseForFiles(sdefname, sdefext, sdefdir, stitle, 0)

    var excelFile = Context.getExcelReader(files[0].getData())

    return excelFile;
}