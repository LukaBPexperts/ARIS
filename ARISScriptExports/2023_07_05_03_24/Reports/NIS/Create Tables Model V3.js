//******************************************************************************
//U koloni KEY ako podatak nije PK ili FK mora biti 1
//******************************************************************************
var nLocale = Context.getSelectedLanguage()
//var selectedDatabases = ArisData.getSelectedDatabases()
var selectedGroups = ArisData.getSelectedGroups()

var selectedGroup = selectedGroups[0]

//var database = selectedDatabases[0]

var excelFile = getExcelFile() //get excell file
var sheets = excelFile.getSheets() //get all sheets

//idi kroz sheet-ove i uzmi sheet-ove prema nazivima
for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() == "Polja") {
        var sheetFields = sheets[i]
    }
    if (sheets[i].getName() == "Tabele") {
        var sheetTables = sheets[i]
    }
    if (sheets[i].getName() == "Veze izmedju tabela") {
        var sheetRelations = sheets[i]
    }
    if (sheets[i].getName() == "Interfejsovi") {
        var sheetInterfaces = sheets[i]
    }
}

//******************************
//**KREIRAJ MODELE SA POLJIMA***
//******************************

//get number of rows in excel table
var rowCount = 0
while (sheetTables.getCell(rowCount, 0) != null) {
    rowCount++
}

//kreiraj model i prvu tablicu
//POCETNE KOORDINATE OBJEKATA - OVO MIJENJAJ ZA SPUSTANJE I POMICANJE OBJEKATA (MODELA) ZBOG ZAGLAVLJA
var CX = 175 //get x coordinate
var CY = 525 //get y coodridante  CY povecavaj za 300 tako da skace u novi red funkcija svaka

//kreiraj model sa tablicom i unutra ovi objekt tablice
//var appName = "GeoBank" //****TU UNESI IME APLIKACIJE ILI MODELA  
var appName = Dialogs.InputBox("Unesi ime aplikacije", "Name input", "no name") //izbaci dialog
var ieModel = selectedGroup.CreateModel(Constants.MT_IEF_DATA_MDL, appName, nLocale) //kreiraj novi model u radnoj grupi 

//idi kroz sheet sa tablicama i kreiraj sve objekte tablica
for (var i = 1; i < rowCount; i++) {
    
    Context.writeStatus("Kreiranje objekta" + i +" od " + rowCount)
    
    tableObj = sheetTables.getCell(i, 0).getCellValue()
    //tableDesc = sheetTables.getCell(i, 1).getCellValue()
    
    try {
        tableDesc = sheetTables.getCell(i, 1).getCellValue()
    }
    catch (ex) { // non-standard
        tableDesc = ""
    }
    
    try {
        tableSchema = sheetTables.getCell(i, 2).getCellValue()
    }
    catch (ex) { // non-standard
        tableSchema = ""
    }
    
    
    //tableSchema = sheetTables.getCell(i, 2).getCellValue()


    var tableObjDef = selectedGroup.CreateObjDef(Constants.OT_ENT_TYPE, tableObj, nLocale) //kreiraj objDef objekta tablice
    var tableObjOcc = ieModel.createObjOcc(Constants.ST_ENT_TYPE, tableObjDef, CX, CY, true, false) //kreiraj objOcc objekta tablice u modelu

    //upisi opisni atribut i schemu po potrebi
    tableObjDef.Attribute(Constants.AT_DESC, nLocale).setValue(tableDesc)
    
    //upisi att schema
    tableObjDef.Attribute(Constants.AT_CSC_ID, nLocale).setValue(tableSchema)

    CY = CY + 200 //povcaj Y koordinatu za 200, znaci spusti se prema dolje za jednu liniju  
    
    
    
    ArisData.getActiveDatabase().clearCaches()
}


//POVEZI OBJEKTE TABLICA

var rowCount = 0
while (sheetRelations.getCell(rowCount, 0) != null) {
    rowCount++
}

//idi kroz sheet sa vezama

//var objOccList = ieModel.ObjOccListFilter()

//ObjOccListFilter ( String value, int localeId, int typeNum, int attrTypeNum )

for (var i = 1; i < rowCount; i++) {

    Context.writeStatus("Kreiranje veze " + i +" od " + rowCount)
        
    tableIn = sheetRelations.getCell(i, 0).getCellValue()
    tableOut = sheetRelations.getCell(i, 1).getCellValue()
    //relationType = sheetRelations.getCell(i, 2).getCellValue()
   // keyName = sheetRelations.getCell(i, 3).getCellValue()
    
    var inObjects = ieModel.ObjOccListFilter(tableIn, nLocale, -1, 1)
    var outObjects = ieModel.ObjOccListFilter(tableOut, nLocale, -1, 1)
    
    

    //SPOJI OBJEKTE TABLICA
    var sourcePoint       
    var targetPoint                
    var import_custom = JavaImporter(Packages.java.awt); // Java Package Import		        

    var sourceObject = inObjects[0] //att objekt
    var targetObject = outObjects[0] //tablica objekt

    with(import_custom) {            

        targetPoint = new Point(targetObject.X() + targetObject.Width() / 2, targetObject.Y() + targetObject.Height())

        sourcePoint = new Point(sourceObject.X(), sourceObject.Y() + sourceObject.Height() / 2)  

    }                
    var points = new Array()        
    points[0] = sourcePoint        
    points[1] = targetPoint

    var newConnection = ieModel.CreateCxnOcc(sourceObject, targetObject, Constants.CT_HAS_REL_WITH, points, false, true).applyTemplate() //kreiraj vezu
    
    
    //UBACI NA VEZU ATTRIBUTE KOJE TREBA UBACITI !!***
    
    
    
    
    
   ArisData.getActiveDatabase().clearCaches()
}

//poslozi modele

Context.writeStatus("Primjenjujem layout")

var layoutValue = ieModel.doLayout()

function getExcelFile() {
    var sdefname = ""
    sdefext = __toString("*.xls!!Excel|*.xls; *.xlsx||")
    var sdefdir = ""
    var stitle = "Chose Excel File"

    var files = Dialogs.BrowseForFiles(sdefname, sdefext, sdefdir, stitle, 0)

    var excelFile = Context.getExcelReader(files[0].getData())

    return excelFile;
}