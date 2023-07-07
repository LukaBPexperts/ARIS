var nLocale = Context.getSelectedLanguage()
var selectedDb = ArisData.getSelectedDatabases()

var excelFile = getExcelFile() //get excell file
var sheets = excelFile.getSheets() //get all sheets



    //get number of rows in excel table
    var rowCount = 0
    while (sheets[0].getCell(rowCount, 0) != null) {
        rowCount++
    }

    //get marked models
    for (var i = 1; i < rowCount; i++) {
        Context.writeStatus("1. Reading excel file row " + (i + 1) + "/" + rowCount)
        var guidTemp = sheets[i].getCell(i, 0).getCellValue() //get GUID from excell       
        var guidNext = sheets[i+1].getCell(i, 0).getCellValue() //get GUID from excell    
              
        var riskDef = selectedDb[0].FindGUID(guidTemp)
        if(riskDef.IsValid()){     
       
           //create objects in FAD connected to risk foer each row, or update existing 
           
           //1. get GUID Mere
           var guidMere = sheets[i].getCell(i, 0).getCellValue() //get GUID from excell
           var meraDef = selectedDb[0].FindGUID(guidTemp)
           
           //if guidMere not empty (meraDef.IsValid()
                //2a. search GUID objDef in Aris DB, and update attributes on that object
            if(meraDef.IsValid()){
            
                //meraDef update attributes
            
            
            } else{
              
           //else, new object 
                //2b. create obj def in db, update attributes on that obj, create occurence in Risk fad, create connection occurence, apply template
                //var startAppObjDef = selectedGroup.CreateObjDef(Constants.OT_APPL_SYS_TYPE, startAppName, nLocale) //kreiraj objDef objekta aplikacije pocetne
                //var startAppObjOcc = model.createObjOcc(Constants.ST_APPL_SYS_TYPE, startAppObjDef, CX, CY, true, false) //kreiraj objOcc objekta pocetne aplikacije
          
           
           
           
           }    
         
         
         //add attributes on risk object
         
        }
                
//####################################################################################################################################
                /* //FROM MATIAS
                g_ntranspoint[0] = new PublicInterface.Point(oSourceOcc.X(),oSourceOcc.Y());
                g_ntranspoint[1] = new PublicInterface.Point(oTargetOcc.X(),oTargetOcc.Y());
                p_cxnLibraryModel.CreateCxnOcc(oSourceOcc, oTargetOcc, currCxnDef, g_ntranspoint, false); 
                */
           
           /*
           //SPAJANJE OBJEKATA NIS ANTE
                        
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
            */
//#####################################################################################################################################
          
          var a =2
            
        
        
        //update risk attributes, read last row i
        

       
    }









function getExcelFile() {
    var sdefname = ""
    sdefext = __toString("*.xls!!Excel|*.xls; *.xlsx||")
    var sdefdir = ""
    var stitle = "Chose excel file"

    var files = Dialogs.BrowseForFiles(sdefname, sdefext, sdefdir, stitle, 0)

    var excelFile = Context.getExcelReader(files[0].getData())

    return excelFile;
}