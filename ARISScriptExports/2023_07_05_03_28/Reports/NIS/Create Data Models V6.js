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
while (sheetFields.getCell(rowCount, 0) != null) {
    rowCount++
}


//dohvati imena tablica iz sheet-a sa podacima
var tablesList = new Array()
for (var i = 1; i < rowCount; i++) {
    //var attElement = sheets[0].getCell(j, 0).getCellValue()
    tablesList.push(sheetFields.getCell(i, 0).getCellValue())
}
tablesList = ArisData.Unique(tablesList) //ostavi samo unique imena tablica

//idi kroz sva imena tablica, za svaku kreiraj model i popuni ga sa atributima
for (var i = 0; i < tablesList.length; i++) {
    
    
    
    //POCETNE KOORDINATE OBJEKATA - OVO MIJENJAJ ZA SPUSTANJE I POMICANJE OBJEKATA (MODELA) ZBOG ZAGLAVLJA
    var CX = 175 //get x coordinate
    var CY = 525 //get y coodridante  CY povecavaj za 300 tako da skace u novi red funkcija svaka

    //kreiraj model sa tablicom i unutra ovi objekt tablice
    var ieModel = selectedGroup.CreateModel(Constants.MT_EERM_ATTR_ALLOC_DGM, tablesList[i], nLocale) //kreiraj novi model u radnoj grupi 
    var tableObjDef = selectedGroup.CreateObjDef(Constants.OT_ENT_TYPE, tablesList[i], nLocale) //kreiraj objDef objekta tablice
    var tableObjOcc = ieModel.createObjOcc(Constants.ST_ENT_TYPE, tableObjDef, CX, CY, true, false) //kreiraj objOcc objekta tablice u modelu

    CY = CY + 200 //povcaj Y koordinatu za 200, znaci spusti se prema dolje za jednu liniju

    //idi kroz sve redove u file-u (atribute), ako je tablica pokraj atributa ista kao ova gore uzeta u obzir kreiraj i spoji objekt attributa na nju

    var attElementPK = new Array()
    var attElementFK = new Array()
    var attElementELSE = new Array()
    var attElementPK_Sym = new Array()
    var attElementFK_Sym = new Array()
    var attElementELSE_Sym = new Array()
    
    var attOpisPK = new Array()
    var attTipPK = new Array()
    var attDuzinaPK = new Array()
    var attDecimalaPK = new Array()
    var attNulablePK = new Array()
    var attDefaultPK = new Array()
    var attIzvornaTabelaPK = new Array()
    var attIzvornoPoljePK = new Array()
    var attNazivStranogKljucaPK = new Array()
    
    var attOpisFK = new Array()
    var attTipFK = new Array()
    var attDuzinaFK = new Array()
    var attDecimalaFK = new Array()
    var attNulableFK = new Array()
    var attDefaultFK = new Array()
    var attIzvornaTabelaFK = new Array()
    var attIzvornoPoljeFK = new Array()
    var attNazivStranogKljucaFK = new Array()
    
    var attOpisELSE = new Array()
    var attTipELSE = new Array()
    var attDuzinaELSE = new Array()
    var attDecimalaELSE = new Array()
    var attNulableELSE = new Array()
    var attDefaultELSE = new Array()
    var attIzvornaTabelaELSE = new Array()
    var attIzvornoPoljeELSE = new Array()
    var attNazivStranogKljucaELSE = new Array()
    
    
    
    for (var j = 1; j < rowCount; j++) {

        var table = sheetFields.getCell(j, 0).getCellValue()

        if (table.toString() == tablesList[i].toString()) {

            //dohvati prvo PK, pa onda FK, pa onda ostale
            if (sheetFields.getCell(j, 8).getCellValue() == "PK") {
                attElementPK.push(sheetFields.getCell(j, 1).getCellValue())
                attElementPK_Sym.push(Constants.ST_KEY_ATTR)
                
                //attOpisPK.push(sheetFields.getCell(j, 2).getCellValue())
                try {
                    attOpisPK.push(sheetFields.getCell(j, 2).getCellValue())
                }
                catch (ex) { // non-standard
                    attOpisPK.push("")
                }
    
                
               // attTipPK.push(sheetFields.getCell(j, 3).getCellValue())
                try {
                    attTipPK.push(sheetFields.getCell(j, 3).getCellValue())
                }
                catch (ex) { // non-standard
                    attTipPK.push("")
                }
                
                //attDuzinaPK.push(sheetFields.getCell(j, 4).getCellValue())
                try {
                    attDuzinaPK.push(sheetFields.getCell(j, 4).getCellValue())
                }
                catch (ex) { // non-standard
                    attDuzinaPK.push("")
                }
                
                //attDecimalaPK.push(sheetFields.getCell(j, 5).getCellValue())
                try {
                    attDecimalaPK.push(sheetFields.getCell(j, 5).getCellValue())
                }
                catch (ex) { // non-standard
                    attDecimalaPK.push("")
                }
                
                //attNulablePK.push(sheetFields.getCell(j, 6).getCellValue())
                try {
                    attNulablePK.push(sheetFields.getCell(j, 6).getCellValue())
                }
                catch (ex) { // non-standard
                    attNulablePK.push("")
                }
                
                //attDefaultPK.push(sheetFields.getCell(j, 7).getCellValue())
                 try {
                    attDefaultPK.push(sheetFields.getCell(j, 7).getCellValue())
                }
                catch (ex) { // non-standard
                    attDefaultPK.push("")
                }
                
                //attIzvornaTabelaPK.push(sheetFields.getCell(j, 9).getCellValue())
                try {
                    attIzvornaTabelaPK.push(sheetFields.getCell(j, 9).getCellValue())
                }
                catch (ex) { // non-standard
                    attIzvornaTabelaPK.push("")
                }
                
                //attIzvornoPoljePK.push(sheetFields.getCell(j, 10).getCellValue())
                try {
                    attIzvornoPoljePK.push(sheetFields.getCell(j, 10).getCellValue())
                }
                catch (ex) { // non-standard
                    attIzvornoPoljePK.push("")
                }
                
                //attNazivStranogKljucaPK.push(sheetFields.getCell(j, 11).getCellValue())
                try {
                    attNazivStranogKljucaPK.push(sheetFields.getCell(j, 11).getCellValue())
                }
                catch (ex) { // non-standard
                    attNazivStranogKljucaPK.push("")
                }
                
                
            }
            else if (sheetFields.getCell(j, 8).getCellValue() == "FK") {
                attElementFK.push(sheetFields.getCell(j, 1).getCellValue())
                attElementFK_Sym.push(Constants.ST_FRGN_KEY_ATTR)
                
                /* attOpisFK.push(sheetFields.getCell(j, 2).getCellValue())
                attTipFK.push(sheetFields.getCell(j, 3).getCellValue())
                attDuzinaFK.push(sheetFields.getCell(j, 4).getCellValue())
                attDecimalaFK.push(sheetFields.getCell(j, 5).getCellValue())
                //attNulableFK.push(sheetFields.getCell(j, 6).getCellValue())
                //attDefaultFK.push(sheetFields.getCell(j, 7).getCellValue())
                //attIzvornaTabelaFK.push(sheetFields.getCell(j, 9).getCellValue())
                //attIzvornoPoljeFK.push(sheetFields.getCell(j, 10).getCellValue())
                //attNazivStranogKljucaFK.push(sheetFields.getCell(j, 11).getCellValue()) */
                
                //attOpisPK.push(sheetFields.getCell(j, 2).getCellValue())
                try {
                    attOpisFK.push(sheetFields.getCell(j, 2).getCellValue())
                }
                catch (ex) { // non-standard
                    attOpisFK.push("")
                }
    
                
               // attTipPK.push(sheetFields.getCell(j, 3).getCellValue())
                try {
                    attTipFK.push(sheetFields.getCell(j, 3).getCellValue())
                }
                catch (ex) { // non-standard
                    attTipFK.push("")
                }
                
                //attDuzinaPK.push(sheetFields.getCell(j, 4).getCellValue())
                try {
                    attDuzinaFK.push(sheetFields.getCell(j, 4).getCellValue())
                }
                catch (ex) { // non-standard
                    attDuzinaFK.push("")
                }            
                
                //attDecimalaPK.push(sheetFields.getCell(j, 5).getCellValue())
                try {
                    attDecimalaFK.push(sheetFields.getCell(j, 5).getCellValue())
                }
                catch (ex) { // non-standard
                    attDecimalaFK.push("")
                }
                
                //attNulablePK.push(sheetFields.getCell(j, 6).getCellValue())
                try {
                    attNulableFK.push(sheetFields.getCell(j, 6).getCellValue())
                }
                catch (ex) { // non-standard
                    attNulableFK.push("")
                }
                
                //attDefaultPK.push(sheetFields.getCell(j, 7).getCellValue())
                 try {
                    attDefaultFK.push(sheetFields.getCell(j, 7).getCellValue())
                }
                catch (ex) { // non-standard
                    attDefaultFK.push("")
                }
                
                //attIzvornaTabelaPK.push(sheetFields.getCell(j, 9).getCellValue())
                try {
                    attIzvornaTabelaFK.push(sheetFields.getCell(j, 9).getCellValue())
                }
                catch (ex) { // non-standard
                    attIzvornaTabelaFK.push("")
                }
                
                //attIzvornoPoljePK.push(sheetFields.getCell(j, 10).getCellValue())
                try {
                    attIzvornoPoljeFK.push(sheetFields.getCell(j, 10).getCellValue())
                }
                catch (ex) { // non-standard
                    attIzvornoPoljeFK.push("")
                }
                
                //attNazivStranogKljucaPK.push(sheetFields.getCell(j, 11).getCellValue())
                try {
                    attNazivStranogKljucaFK.push(sheetFields.getCell(j, 11).getCellValue())
                }
                catch (ex) { // non-standard
                    attNazivStranogKljucaFK.push("")
                }
                
            }
            else {
                attElementELSE.push(sheetFields.getCell(j, 1).getCellValue())
                attElementELSE_Sym.push(Constants.ST_DESC_ATTR)
                
                /* attOpisELSE.push(sheetFields.getCell(j, 2).getCellValue())
                attTipELSE.push(sheetFields.getCell(j, 3).getCellValue())
                attDuzinaELSE.push(sheetFields.getCell(j, 4).getCellValue())
                attDecimalaELSE.push(sheetFields.getCell(j, 5).getCellValue())
                //attNulableELSE.push(sheetFields.getCell(j, 6).getCellValue())
                //attDefaultELSE.push(sheetFields.getCell(j, 7).getCellValue())
                //attIzvornaTabelaELSE.push(sheetFields.getCell(j, 9).getCellValue())
                //attIzvornoPoljeELSE.push(sheetFields.getCell(j, 10).getCellValue())
                //attNazivStranogKljucaELSE.push(sheetFields.getCell(j, 11).getCellValue()) */
                
                //attOpisPK.push(sheetFields.getCell(j, 2).getCellValue())
                try {
                    attOpisELSE.push(sheetFields.getCell(j, 2).getCellValue())
                }
                catch (ex) { // non-standard
                    attOpisELSE.push("")
                }
    
                
               // attTipPK.push(sheetFields.getCell(j, 3).getCellValue())
                try {
                    attTipELSE.push(sheetFields.getCell(j, 3).getCellValue())
                }
                catch (ex) { // non-standard
                    attTipELSE.push("")
                }         
                
                //attDuzinaPK.push(sheetFields.getCell(j, 4).getCellValue())
                try {
                    attDuzinaELSE.push(sheetFields.getCell(j, 4).getCellValue())
                }
                catch (ex) { // non-standard
                    attDuzinaELSE.push("")
                }            
                
                //attDecimalaPK.push(sheetFields.getCell(j, 5).getCellValue())
                try {
                    attDecimalaELSE.push(sheetFields.getCell(j, 5).getCellValue())
                }
                catch (ex) { // non-standard
                    attDecimalaELSE.push("")
                }
                
                //attNulablePK.push(sheetFields.getCell(j, 6).getCellValue())
                try {
                    attNulableELSE.push(sheetFields.getCell(j, 6).getCellValue())
                }
                catch (ex) { // non-standard
                    attNulableELSE.push("")
                }
                
                //attDefaultPK.push(sheetFields.getCell(j, 7).getCellValue())
                 try {
                    attDefaultELSE.push(sheetFields.getCell(j, 7).getCellValue())
                }
                catch (ex) { // non-standard
                    attDefaultELSE.push("")
                }
                
                //attIzvornaTabelaPK.push(sheetFields.getCell(j, 9).getCellValue())
                try {
                    attIzvornaTabelaELSE.push(sheetFields.getCell(j, 9).getCellValue())
                }
                catch (ex) { // non-standard
                    attIzvornaTabelaELSE.push("")
                }
                
                //attIzvornoPoljePK.push(sheetFields.getCell(j, 10).getCellValue())
                try {
                    attIzvornoPoljeELSE.push(sheetFields.getCell(j, 10).getCellValue())
                }
                catch (ex) { // non-standard
                    attIzvornoPoljeELSE.push("")
                }
                
                //attNazivStranogKljucaPK.push(sheetFields.getCell(j, 11).getCellValue())
                try {
                    attNazivStranogKljucaELSE.push(sheetFields.getCell(j, 11).getCellValue())
                }
                catch (ex) { // non-standard
                    attNazivStranogKljucaELSE.push("")
                }
                
                
            }
        }
    }

    //sada kreiraj jedan niz gdje ce na prvom mistu biti PK, pa FK, pa ostali        
    var attElement = attElementPK.concat(attElementFK, attElementELSE)
    var attElement_Sym = attElementPK_Sym.concat(attElementFK_Sym, attElementELSE_Sym)
    
    var attOpis = attOpisPK.concat(attOpisFK, attOpisELSE)
    var attTip = attTipPK.concat(attTipFK,attTipELSE)
    var attDuzina = attDuzinaPK.concat(attDuzinaFK,attDuzinaELSE)
    var attDecimala = attDecimalaPK.concat(attDecimalaFK,attDecimalaELSE)
    var attNulable = attNulablePK.concat(attNulableFK,attNulableELSE)   
    var attDefault = attDefaultPK.concat(attDefaultFK,attDefaultELSE)
    var attIzvornaTabela = attIzvornaTabelaPK.concat(attIzvornaTabelaFK,attIzvornaTabelaELSE)
    var attIzvornoPolje = attIzvornoPoljePK.concat(attIzvornoPoljeFK,attIzvornoPoljeELSE)
    var attNazivStranogKljuca = attNazivStranogKljucaPK.concat(attNazivStranogKljucaFK,attNazivStranogKljucaELSE)

    //sada idi kroz sve objekte po redu i kreiraj ih
    for (var j = 0; j < attElement.length; j++) {
    
    Context.writeStatus("Model "+ i + " od "+ tablesList.length + " Kreiranje objekta " + j + " od " + attElement.length)
        
        //kreiraj objekt u modelu i spoji ga

        //kreiraj objDef attributa
        var attObjDef = selectedGroup.CreateObjDef(Constants.OT_ERM_ATTR, attElement[j], nLocale)
        
        //UPISI SVE POTREBNE ATTRIBUTE U OBJEKT

         attObjDef.Attribute(Constants.AT_DESC, nLocale).setValue(attOpis[j])
        
         attObjDef.Attribute(Constants.AT_TYPE, nLocale).setValue(attTip[j])
        
        try{
         attObjDef.Attribute(Constants.AT_LEN_1, nLocale).setValue(parseInt(attDuzina[j]))
        }
        catch (ex){
            
        }
        
        try{
            attObjDef.Attribute(Constants.AT_DEC_DIGIT, nLocale).setValue(parseInt(attDecimala[j]))
        }
        catch(ex){
            
        }
        
        //attNulable[j]
         if(attNulable[j] == "Y"){
         attObjDef.Attribute(Constants.AT_OPT_ATTR, nLocale).setValue(true)
         }
         else if(attNulable[j] == "N"){
             attObjDef.Attribute(Constants.AT_OPT_ATTR, nLocale).setValue(false)
         }
         
         
         attObjDef.Attribute(Constants.AT_DEFAULT_VALUE, nLocale).setValue(attDefault[j])
         
         //attIzvornaTabela[j]         
         //attIzvornoPolje[j]
         //attNazivStranogKljuca[j]
        
        //kreiraj ObjOcc attributa u modelu, pazi na simbol PK, FK, ELSE
        //objekte atributa ispisi uvijek na X kordinati 525, Y koordinatu pomiči za 200

        var attObjOcc = ieModel.createObjOcc(attElement_Sym[j], attObjDef, 525, CY, true, false)
        CY = CY + 200

        //spoji objekte
        var sourcePoint       
        var targetPoint                
        var import_custom = JavaImporter(Packages.java.awt); // Java Package Import		        

        var sourceObject = attObjOcc //att objekt
        var targetObject = tableObjOcc //tablica objekt

        with(import_custom) {            
            /* targetPoint = new Point(sourceObject.X() + sourceObject.Width() / 2, sourceObject.Y() + sourceObject.Height())
            sourcePoint = new Point(targetObject.X() + targetObject.Width() / 2, targetObject.Y())    */

            targetPoint = new Point(targetObject.X() + targetObject.Width() / 2, targetObject.Y() + targetObject.Height())
            //sourcePoint = new Point(sourceObject.X() + sourceObject.Width() / 2, sourceObject.Y())  

            sourcePoint = new Point(sourceObject.X(), sourceObject.Y() + sourceObject.Height() / 2)  

            middlePoint = new Point(sourceObject.X() - 125, sourceObject.Y() + sourceObject.Height() / 2)  //točka lomljena veze

        }                
        var points = new Array()        
        points[0] = sourcePoint        
        points[1] = targetPoint
        points[2] = middlePoint

        //stvori vezu izmedju trenutno kreiranog objekta i onog objekta koji se nalazi prije njega
        //provjeri tip veze prema tipu objOcc kojeg crtas !!!

        if (attElement_Sym[j] == Constants.ST_KEY_ATTR) {
            var newConnection = ieModel.CreateCxnOcc(sourceObject, targetObject, Constants.CT_IS_PRIM_KEY_FOR_1, points, false, true).applyTemplate() //kreiraj vezu
        }
        else if (attElement_Sym[j] == Constants.ST_FRGN_KEY_ATTR) {
            var newConnection = ieModel.CreateCxnOcc(sourceObject, targetObject, Constants.CT_IS_FRGN_KEY_FOR_1, points, false, true).applyTemplate() //kreiraj vezu
        }
        else if (attElement_Sym[j] == Constants.ST_DESC_ATTR) {
            var newConnection = ieModel.CreateCxnOcc(sourceObject, targetObject, Constants.CT_IS_DESC_FOR_1, points, false, true).applyTemplate() //kreiraj vezu
        }
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