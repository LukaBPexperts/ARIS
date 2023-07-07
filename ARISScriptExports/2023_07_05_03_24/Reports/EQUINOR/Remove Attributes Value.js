var languageRemove = 1030 //Danish
var outputFile = Context.createOutputObject()
var nLocale = Context.getSelectedLanguage()

var selectedDatabases = ArisData.getSelectedDatabases()
var selectedDB = selectedDatabases[0]

//var modelType = []

//var groups = selectedDB.Find(Constants.SEARCH_GROUP)
Context.writeStatus("Reading models")
var models = selectedDB.Find(Constants.SEARCH_MODEL)
Context.writeStatus("Reading objects")
var objects = selectedDB.Find(Constants.SEARCH_OBJDEF)

var items = models.concat(objects)

//var groups = null;
var models = null;
var objects = null;

for (var i = 0; i < items.length; i++) {
    Context.writeStatus("Working on model/object: " + (i + 1) + '/' + (items.length + 1))
    removeAttributes(items[i], languageRemove)
    items[i] = null; //clear array
}
items = null;

outputFile.WriteReport();

function removeAttributes(item, language) {
    var attList = item.AttrList(language, false)
    for (var i = 0; i < attList.length; i++) {
        if (attList[i].LanguageDependence() != 1) {
            var value = attList[i].Delete()
            if (value) {
                //log deleted
                outputFile.OutputTxt("DELETED | " + item.Type() + ": " + item.Name(nLocale) + " (" + item.GUID() + ") " + " Attribute: " + attList[i].Type() + "\n");
            }
            else {
                //
                outputFile.OutputTxt("ERROR | " + item.Type() + ": " + item.Name(nLocale) + " (" + item.GUID() + ") " + " Attribute: " + attList[i].Type() + " Unable to delete attribute" + "\n");
            }
        }
    }
}