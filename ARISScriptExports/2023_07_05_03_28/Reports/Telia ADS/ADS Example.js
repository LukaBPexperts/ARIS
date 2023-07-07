//var outputFile = Context.createOutputObject()
//outputFile.OutputTxt("Aris Document Storage test");
//outputFile.WriteReport();
//writeDocumentADS() 

var byteDoc = readDocumentADS()

//var output = Context.getOutputObject("test.xlsx", byteDoc)

var output = Context.createExcelWorkbook("testFile.xlsx", byteDoc )


//var workbookReader = Context.getExcelReader(byteDoc)

output.write()



function writeDocumentADS() {
    var adsReport = Context.getComponent("ADS")
    var repository = adsReport.getADSRepository("portal")
    var rootFolder = repository.getRootFolder();

    var folderName = "ADS Test\\"
    var adsFolder = repository.getFolder(rootFolder, folderName);

    if (!adsFolder)
        adsFolder = repository.createFolder(rootFolder, folderName);

    var docInfo = repository.createDocumentMetaInfo("testFile", "testFile.txt", "Test File");
    var contentByteArray = Context.getFile(Context.getSelectedFile(), Constants.LOCATION_OUTPUT)
    var document = repository.createAndOverwriteExistingDocument(adsFolder, docInfo, new java.io.ByteArrayInputStream(contentByteArray))
}

function readDocumentADS() {
    var adsReport = Context.getComponent("ADS")
    var repository = adsReport.getADSRepository("portal")
    var rootFolder = repository.getRootFolder();

    var folderName = "XLS Template\\"
    var adsFolder = repository.getFolder(rootFolder, folderName);
    
    var fileName = "BPMN_2_input_data_entry-page.xlsx"
    var documentTemplate = repository.getDocument( adsFolder, null, fileName)
    
    var docContent = documentTemplate.getDocumentContent()
    
    var byteDoc = Packages.com.google.common.io.ByteStreams.toByteArray(docContent) //convert InputStream to Byte Array
    
    return byteDoc;
}