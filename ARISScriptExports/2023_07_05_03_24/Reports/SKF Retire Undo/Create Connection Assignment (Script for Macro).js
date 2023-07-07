//*******************************************************************************************************************************************
//******************************************************CUSTOMIZE****************************************************************************
//*******************************************************************************************************************************************
var adminUsername = "customized_script_runner" //user must have access to the models which we want to assign
var adminPassword = "CSR"
//*******************************************************************************************************************************************
//*******************************************************************************************************************************************

//**********************************************************************************************************************
//**************************************DON'T CHANGE ANYTHING BELOW THIS LINE********************************************
//**********************************************************************************************************************
var nLocale = Context.getSelectedLanguage()
var outputFile = Context.createOutputObject()
var entireMethodFilterGUID = "dd838074-ac29-11d4-85b8-00005a4053ff"

var cxnGUID = Context.getProperty("cxnGUID")
var modelGUID = Context.getProperty("modelGUID")
var dbName = Context.getProperty("dbName")

var dbActive = ArisData.openDatabase(dbName, adminUsername, adminPassword, entireMethodFilterGUID, Context.getSelectedLanguage(), false);

var model = dbActive.FindGUID(modelGUID)
var cxnDef = dbActive.FindGUID(cxnGUID)

var assignmentInfo = cxnDef.CreateAssignment(model)

Context.setProperty("assignmentInfo", assignmentInfo)
Context.setProperty("cxnType", cxnDef.ActiveType())
Context.setProperty("fromObject", cxnDef.SourceObjDef().Name(nLocale))
Context.setProperty("toObject", cxnDef.TargetObjDef().Name(nLocale))

//outputFile.OutputTxt("assignmentInfo");

//outputFile.WriteReport();