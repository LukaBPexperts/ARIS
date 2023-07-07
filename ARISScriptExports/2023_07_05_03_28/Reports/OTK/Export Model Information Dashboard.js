var oOutput = Context.createOutputObject()
var nLocale = Context.getSelectedLanguage()
var startDatabase = ArisData.getSelectedDatabases()

//80caed51-cf77-11e7-6a8c-005056be329d FindGUID ( String sGUID )
var startGroup = startDatabase[0].FindGUID("80caed51-cf77-11e7-6a8c-005056be329d")
var modelsList = startGroup.ModelList(true, null)

//var startGroup = ArisData.getSelectedGroups()
//var modelsList = startGroup[0].ModelList(true, null)

main()

oOutput.WriteReport()

function main() {
    oOutput.BeginTable(100, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
    createTableHeader()

    //go trought models
    for (var i = 0; i < modelsList.length; i++) {
        if (modelsList[i].Type() != "Function allocation diagram") {

            var path = modelsList[i].Group().Path(nLocale)
            var modelName = modelsList[i].Name(nLocale)
            var modelStatus = modelsList[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("0f154371-04df-11e8-6a8c-005056be329d"), nLocale).GetValue(false)
            var responsibleModeller = modelsList[i].Attribute(Constants.AT_PERS_RESP, nLocale).GetValue(false)
            var lastChange = modelsList[i].Attribute(Constants.AT_LAST_CHNG_2, nLocale).GetValue(false)
            var type = modelsList[i].Type()
            var timeOfGeneration = modelsList[i].Attribute(Constants.AT_CREAT_TIME_STMP, nLocale).GetValue(false)
            var owner = modelsList[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("e8f82d70-df17-11e7-6a8c-005056be329d"), nLocale).GetValue(false)
            var ownerName = modelsList[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("d1427391-d5ff-11e8-6eb2-005056011ee4"), nLocale).GetValue(false)
            var lastReview = modelsList[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("b41a8c11-df17-11e7-6a8c-005056be329d"), nLocale).GetValue(false)

            createTableRow(path, modelName, modelStatus, responsibleModeller, lastChange, type, timeOfGeneration, owner, ownerName)
        }
    }
    
    oOutput.EndTable("Sheet 1", 100, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
}

function createTableRow(path, modelName, modelStatus, responsibleModeller, lastChange, type, timeOfGeneration, owner, ownerName, lastReview) {
    oOutput.TableRow()
    oOutput.TableCell(path, 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
    oOutput.TableCell(modelName, 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
    oOutput.TableCell(modelStatus, 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
    oOutput.TableCell(responsibleModeller, 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
    oOutput.TableCell(lastChange, 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
    oOutput.TableCell(type, 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
    oOutput.TableCell(timeOfGeneration, 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
    oOutput.TableCell(owner, 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
    oOutput.TableCell(ownerName, 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
    oOutput.TableCell(lastReview, 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
}

function createTableHeader() {
    oOutput.TableRow()
    oOutput.TableCell("path", 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
    oOutput.TableCell("model", 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
    oOutput.TableCell("model_status", 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
    oOutput.TableCell("responsible_modeller", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
    oOutput.TableCell("last_change", 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
    oOutput.TableCell("type", 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
    oOutput.TableCell("time_of_generation", 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
    oOutput.TableCell("owner", 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
    oOutput.TableCell("owner_name", 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
    oOutput.TableCell("lastReview", 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)

}