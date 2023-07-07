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
while (sheetInterfaces.getCell(rowCount, 0) != null) {
    rowCount++
}


//dohvati imena aplikacija iz sheet-a interfejsovi
var appList = new Array()
for (var i = 1; i < rowCount; i++) {
    //var attElement = sheets[0].getCell(j, 0).getCellValue()
    appList.push(sheetInterfaces.getCell(i, 3).getCellValue())
}

appList = ArisData.Unique(appList) //ostavi samo unique imena tablica

var startAppName = Dialogs.InputBox("Unesi ime startne aplikacije", "Name input", "no name") //izbaci dialog
//var startAppName = "OIS"
var startAppObjDef = selectedGroup.CreateObjDef(Constants.OT_APPL_SYS_TYPE, startAppName, nLocale) //kreiraj objDef objekta aplikacije pocetne

//idi kroz sva imena tablica, za svaku kreiraj model i popuni ga sa atributima
for (var i = 0; i < appList.length; i++) {

    //POCETNE KOORDINATE OBJEKATA - OVO MIJENJAJ ZA SPUSTANJE I POMICANJE OBJEKATA (MODELA) ZBOG ZAGLAVLJA
    var CX = 175 //get x coordinate
    var CY = 525 //get y coodridante  CY povecavaj za 300 tako da skace u novi red funkcija svaka

    //kreiraj model sa tablicom i unutra ovi objekt tablice
    var model = selectedGroup.CreateModel(Constants.MT_PRG_STRCT_CHRT, startAppName + " - " + appList[i], nLocale) //kreiraj novi model u radnoj grupi 

    var startAppObjOcc = model.createObjOcc(Constants.ST_APPL_SYS_TYPE, startAppObjDef, CX, CY, true, false) //kreiraj objOcc objekta pocetne aplikacije

    var inAttOccList = new Array()
    for (var j = 1; j < rowCount; j++) {

        var inputApp = sheetInterfaces.getCell(j, 3).getCellValue()

        if (inputApp.toString() == appList[i].toString()) {

            CY = 925

            var outAttName = sheetInterfaces.getCell(j, 1).getCellValue()
            var inAttName = sheetInterfaces.getCell(j, 4).getCellValue()
            var interfaceName = sheetInterfaces.getCell(j, 0).getCellValue()

            //kreiraj defObjekte od polja in/out, occ objekte out polja spremaj u listu jednu da ih kasnije spojis lako na ovaj izlazni objekt aplikacije
            var outAttObjDef = selectedGroup.CreateObjDef(Constants.OT_ENT_TYPE, outAttName, nLocale) //kreiraj objDef objekta attributa izlaznog
            var inAttObjDef = selectedGroup.CreateObjDef(Constants.OT_ENT_TYPE, inAttName, nLocale) //kreiraj objDef objekta attributa ulaznog

            //kreiraj defObj naziva interfejsa
            var interfaceObjDef = selectedGroup.CreateObjDef(Constants.OT_CLS, interfaceName, nLocale) //kreiraj objDef objekta interfejsa

            //nacrtaj to u modelu i spoji
            var outAttObjOcc = model.createObjOcc(Constants.ST_ENT_TYPE, outAttObjDef, CX, CY, true, false) //kreiraj objOcc objekta pocetne aplikacije

            CY = CY + 400 //povcaj Y koordinatu za 400, znaci spusti se prema dolje za dvije linije barem

            var interfaceObjOcc = model.createObjOcc(Constants.ST_IF, interfaceObjDef, CX, CY, true, false) //kreiraj objOcc objekta pocetne aplikacije

            CY = CY + 400 //povcaj Y koordinatu za 400, znaci spusti se prema dolje za dvije linije barem

            var inAttObjOcc = model.createObjOcc(Constants.ST_ENT_TYPE, inAttObjDef, CX, CY, true, false) //kreiraj objOcc objekta pocetne aplikacije
            inAttOccList.push(inAttObjOcc) //lista za occ izlaznih objekata zbog kasnijeg spajanja na zadnji objekt aplikacije

            //pomakni koordinate desno, X povecaj
            CX = CX + 550

            //SPAJANJE OBJEKATA********
                        
            var import_custom = JavaImporter(Packages.java.awt); // Java Package Import		        

            //********************************************************************************************************
            var sourceObject = startAppObjOcc //att objekt
            var targetObject = outAttObjOcc //tablica objekt

            with(import_custom) {            
                var sourcePoint = new Point(sourceObject.X() + sourceObject.Width() / 2, sourceObject.Y() + sourceObject.Height())  
                var targetPoint = new Point(targetObject.X() + targetObject.Width() / 2, targetObject.Y())
                var middlePoint = new Point(sourceObject.X() + sourceObject.Width() / 2, sourceObject.Y() + sourceObject.Height() + 125)  //točka lomljena veze
                var middlePoint2 = new Point(targetObject.X() + targetObject.Width() / 2, targetObject.Y() - targetObject.Height() + 25)  //točka lomljena veze

            }                
            var points = new Array()        
            points[0] = sourcePoint        
            points[1] = middlePoint
            points[2] = middlePoint2
            points[3] = targetPoint

            //kreiraj vezu izmedju objekta applikacije start i podataka
            var newConnection = model.CreateCxnOcc(sourceObject, targetObject, Constants.CT_HAS_OUT, points, false, true).applyTemplate() //kreiraj vezu

            //******************************************************************************************************
            //NOVI RED VEZA ******
            var sourceObject = outAttObjOcc //att objekt
            var targetObject = interfaceObjOcc //tablica objekt

            with(import_custom) {            
                var sourcePoint = new Point(sourceObject.X() + sourceObject.Width() / 2, sourceObject.Y() + sourceObject.Height())  
                var targetPoint = new Point(targetObject.X() + targetObject.Width() / 2, targetObject.Y())
            }                
            var points = new Array()        
            points[0] = sourcePoint        
            points[1] = targetPoint

            //kreiraj vezu izmedju podataka i interfacea
            var newConnection = model.CreateCxnOcc(sourceObject, targetObject, Constants.CT_IS_INP_FOR, points, false, true).applyTemplate() //kreiraj vezu

            //*************************************************************
            //NOVI RED VEZA ******
            var sourceObject = interfaceObjOcc //att objekt
            var targetObject = inAttObjOcc //tablica objekt

            with(import_custom) {            
                var sourcePoint = new Point(sourceObject.X() + sourceObject.Width() / 2, sourceObject.Y() + sourceObject.Height())  
                var targetPoint = new Point(targetObject.X() + targetObject.Width() / 2, targetObject.Y())
            }                
            var points = new Array()        
            points[0] = sourcePoint        
            points[1] = targetPoint

            //kreiraj vezu izmedju podataka i interfacea
            var newConnection = model.CreateCxnOcc(sourceObject, targetObject, Constants.CT_HAS_OUT, points, false, true).applyTemplate() //kreiraj vezu

            //*****************************************************************************************

        }
    }

    //kada je sve to iscrtano na kraj dodaj jos jedan objekt aplikacije ove appList[i] i spoji sve donje objekte na nju
    var endAppObjDef = selectedGroup.CreateObjDef(Constants.OT_APPL_SYS_TYPE, appList[i], nLocale) //kreiraj objDef objekta aplikacije pocetne

    var endAppObjOcc = model.createObjOcc(Constants.ST_APPL_SYS_TYPE, endAppObjDef, 175, 2125, true, false) //kreiraj objOcc objekta pocetne aplikacije

    for (var j = 0; j < inAttOccList.length; j++) {

        var import_custom = JavaImporter(Packages.java.awt); // Java Package Import	

        var sourceObject = inAttOccList[j] //att objekt
        var targetObject = endAppObjOcc //tablica objekt

        with(import_custom) {                    
            var sourcePoint = new Point(sourceObject.X() + sourceObject.Width() / 2, sourceObject.Y() + sourceObject.Height())  
            var targetPoint = new Point(targetObject.X() + targetObject.Width() / 2, targetObject.Y())
            var middlePoint = new Point(sourceObject.X() + sourceObject.Width() / 2, sourceObject.Y() + sourceObject.Height() + 125)  //točka lomljena veze
            var middlePoint2 = new Point(targetObject.X() + targetObject.Width() / 2, targetObject.Y() - targetObject.Height() + 25)  //točka lomljena veze
        }                
        var points = new Array()        
        points[0] = sourcePoint        
        points[1] = middlePoint
        points[2] = middlePoint2
        points[3] = targetPoint
        
        //kreiraj vezu izmedju objekta applikacije start i podataka
        var newConnection = model.CreateCxnOcc(sourceObject, targetObject, Constants.CT_IS_INP_FOR, points, false, true).applyTemplate() //kreiraj vezu
    }
}

function getExcelFile() {
    var sdefname = ""
    sdefext = __toString("*.xls!!Excel|*.xls; *.xlsx||")
    var sdefdir = ""
    var stitle = "Chose Excel File"

    var files = Dialogs.BrowseForFiles(sdefname, sdefext, sdefdir, stitle, 0)

    var excelFile = Context.getExcelReader(files[0].getData())

    return excelFile;
}