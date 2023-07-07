var oOutput = Context.createOutputObject()
var nLocale = Context.getSelectedLanguage()
var startDatabase = ArisData.getSelectedDatabases()
var xmlOutput = Context.createXMLOutputObject(Context.getSelectedFile(), "Root");

var startGroup = startDatabase[0].FindGUID("80caed51-cf77-11e7-6a8c-005056be329d") //get "2. Processes" group
var modelsList = startGroup.ModelList(true, null)

main()

var adsReport = Context.getComponent("ADS")
var repository = adsReport.getADSRepository("portal")
var rootFolder = repository.getRootFolder();

var folderName = "OTK Data\\"
var adsFolder = repository.getFolder(rootFolder, folderName);

if (!adsFolder)
    adsFolder = repository.createFolder(rootFolder, folderName);

var docInfo = repository.createDocumentMetaInfo("Models data", "modelsData.xml", "Models data");

contentByteArray = xmlOutput.getDocumentContent()
var document = repository.createAndOverwriteExistingDocument(adsFolder, docInfo, new java.io.ByteArrayInputStream(contentByteArray))

function main() {
    for (var i = 0; i < modelsList.length; i++) {
        if (modelsList[i].TypeNum() != Constants.MT_FUNC_ALLOC_DGM) {
            var xmlModelData = writeXmlItem(xmlOutput, xmlOutput.getRootElement(), modelsList[i], "ModelData");
        }
    }
}

function writeXmlItem(xmlOutput, xmlParent, mItem, sElement) {
    var xmlItem = xmlOutput.addElement(xmlParent, sElement);
    var path = mItem.Group().Path(nLocale)
    var modelName = mItem.Name(nLocale)
    var modelStatus = mItem.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("0f154371-04df-11e8-6a8c-005056be329d"), nLocale).GetValue(false)
    var responsibleModeller = mItem.Attribute(Constants.AT_PERS_RESP, nLocale).GetValue(false)
    var lastChange = mItem.Attribute(Constants.AT_LAST_CHNG_2, nLocale).GetValue(false)
    var lastChangeManual = mItem.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("40109581-df17-11e7-6a8c-005056be329d"), nLocale).GetValue(false)  
    var type = mItem.Type()
    var timeOfGeneration = mItem.Attribute(Constants.AT_CREAT_TIME_STMP, nLocale).GetValue(false)
    var owner = mItem.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("e8f82d70-df17-11e7-6a8c-005056be329d"), nLocale).GetValue(false)
    var ownerName = mItem.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("d1427391-d5ff-11e8-6eb2-005056011ee4"), nLocale).GetValue(false)
    var lastReview = mItem.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("b41a8c11-df17-11e7-6a8c-005056be329d"), nLocale).GetValue(false)
    var modelGUID = mItem.GUID()

    xmlItem.setAttribute("path", path);
    xmlItem.setAttribute("model_name", modelName);
    xmlItem.setAttribute("model_status", modelStatus);
    xmlItem.setAttribute("responsible_modeller", responsibleModeller);
    xmlItem.setAttribute("last_change", lastChange);
    xmlItem.setAttribute("type", type);
    xmlItem.setAttribute("time_of_generation", timeOfGeneration);
    xmlItem.setAttribute("owner", owner);
    xmlItem.setAttribute("owner_name", ownerName);
    xmlItem.setAttribute("last_review", lastReview);
    xmlItem.setAttribute("last_change_manual", lastChangeManual);
    xmlItem.setAttribute("model_GUID", modelGUID);

    return xmlItem;
}