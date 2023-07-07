var nLocale = Context.getSelectedLanguage()
var selectedDb = ArisData.getSelectedDatabases()

var gExcelObject = Context.createExcelWorkbook(Context.getSelectedFile());
var gsSheetName = "Printed at ";
var giRowCounter = 0;
var rowColorBr = 1;
var gsBulletSign = "\u2022";
printRowHeaders(gExcelObject);

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
    var guidTemp = sheets[0].getCell(i, 36-1).getCellValue() //get GUID from excell       
    //var guidNext = sheets[0].getCell(i+1, 36).getCellValue() //get GUID from excell    

    var riskDef = selectedDb[0].FindGUID(guidTemp)
    if (riskDef.IsValid()) {

        //create objects in FAD connected to risk foer each row, or update existing 

        //1. get GUID Mere
        var guidMere = sheets[0].getCell(i, 37-1).getCellValue() //get GUID from excell
        var meraDef = selectedDb[0].FindGUID(guidMere)

        //if guidMere not empty (meraDef.IsValid()
        //2a. search GUID objDef in Aris DB, and update attributes on that object
        if (meraDef.IsValid()) {

            //meraDef update attributes 24-29  
            var meraNameTemp = sheets[0].getCell(i, 23-1).getCellValue()
            var meraNameTempModified = meraNameTemp.slice(0,240)              
            var opisOcekivanogRezTemp = sheets[0].getCell(i, 24-1).getCellValue()
            var rokIspunjenaAktivnostiTemp = sheets[0].getCell(i, 25-1).getCellValue()
            var odgovoranTemp = sheets[0].getCell(i, 26-1).getCellValue()
            var troskoviAktivnostiTemp = sheets[0].getCell(i, 27-1).getCellValue()
            var opisRealizacijeMeraTemp = sheets[0].getCell(i, 28-1).getCellValue()
            var statusRealizacijeMeraTemp = sheets[0].getCell(i, 29-1).getCellValue()
            
            var opisOcekivanogRezOld = meraDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("74558751-4e17-11ed-7867-00505682218a"), nLocale).getValue()
            var rokIspunjenaAktivnostiOld = meraDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("88ef6c80-4e17-11ed-7867-00505682218a"), nLocale).getValue()
            var odgovoranOld = meraDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("9580bd51-4e17-11ed-7867-00505682218a"), nLocale).getValue()
            var troskoviAktivnostiOld = meraDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("a983b9b1-4e17-11ed-7867-00505682218a"), nLocale).getValue()
            var opisRealizacijeMeraOld = meraDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("baa763e0-4e17-11ed-7867-00505682218a"), nLocale).getValue()
            var statusRealizacijeMeraOld = meraDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("cb0fd461-4e17-11ed-7867-00505682218a"), nLocale).getValue()
            
            if(opisOcekivanogRezTemp.replaceAll("[\n\r]", "") != opisOcekivanogRezOld.replaceAll("[\n\r]", "")){
                //create LOG                
                printRow(riskDef.Name(nLocale), "", "", "", meraDef.Name(nLocale), meraDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("74558751-4e17-11ed-7867-00505682218a"), nLocale).Type(), opisOcekivanogRezTemp, opisOcekivanogRezOld, gExcelObject, rowColorBr)
                rowColorBr++                  
            }
            if(rokIspunjenaAktivnostiTemp.replaceAll("[\n\r]", "") != rokIspunjenaAktivnostiOld.replaceAll("[\n\r]", "")){
                //create LOG                
                printRow(riskDef.Name(nLocale), "", "", "", meraDef.Name(nLocale), meraDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("88ef6c80-4e17-11ed-7867-00505682218a"), nLocale).Type(), rokIspunjenaAktivnostiTemp, rokIspunjenaAktivnostiOld, gExcelObject, rowColorBr)
                rowColorBr++                  
            }
            if(odgovoranTemp.replaceAll("[\n\r]", "") != odgovoranOld.replaceAll("[\n\r]", "")){
                //create LOG                
                printRow(riskDef.Name(nLocale), "", "", "", meraDef.Name(nLocale), meraDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("9580bd51-4e17-11ed-7867-00505682218a"), nLocale).Type(), odgovoranTemp, odgovoranOld, gExcelObject, rowColorBr)
                rowColorBr++                  
            }
            if(troskoviAktivnostiTemp.replaceAll("[\n\r]", "") != troskoviAktivnostiOld.replaceAll("[\n\r]", "")){
                //create LOG                
                printRow(riskDef.Name(nLocale), "", "", "", meraDef.Name(nLocale), meraDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("a983b9b1-4e17-11ed-7867-00505682218a"), nLocale).Type(), troskoviAktivnostiTemp, troskoviAktivnostiOld, gExcelObject, rowColorBr)
                rowColorBr++                  
            }
            if(opisRealizacijeMeraTemp.replaceAll("[\n\r]", "") != opisRealizacijeMeraOld.replaceAll("[\n\r]", "")){
                //create LOG                
                printRow(riskDef.Name(nLocale), "", "", "", meraDef.Name(nLocale), meraDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("baa763e0-4e17-11ed-7867-00505682218a"), nLocale).Type(), opisRealizacijeMeraTemp, opisRealizacijeMeraOld, gExcelObject, rowColorBr)
                rowColorBr++                  
            }
            if(statusRealizacijeMeraTemp.replaceAll("[\n\r]", "") != statusRealizacijeMeraOld.replaceAll("[\n\r]", "")){
                //create LOG                
                printRow(riskDef.Name(nLocale), "", "", "", meraDef.Name(nLocale), meraDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("cb0fd461-4e17-11ed-7867-00505682218a"), nLocale).Type(), statusRealizacijeMeraTemp, statusRealizacijeMeraOld, gExcelObject, rowColorBr)
                rowColorBr++                  
            } 
                                    
            var opisOcekivanogRez = meraDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("74558751-4e17-11ed-7867-00505682218a"), nLocale).setValue(opisOcekivanogRezTemp)
            var rokIspunjenaAktivnosti = meraDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("88ef6c80-4e17-11ed-7867-00505682218a"), nLocale).setValue(rokIspunjenaAktivnostiTemp)
            var odgovoran = meraDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("9580bd51-4e17-11ed-7867-00505682218a"), nLocale).setValue(odgovoranTemp)
            var troskoviAktivnosti = meraDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("a983b9b1-4e17-11ed-7867-00505682218a"), nLocale).setValue(troskoviAktivnostiTemp)
            var opisRealizacijeMera = meraDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("baa763e0-4e17-11ed-7867-00505682218a"), nLocale).setValue(opisRealizacijeMeraTemp)
            var statusRealizacijeMera = meraDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("cb0fd461-4e17-11ed-7867-00505682218a"), nLocale).setValue(statusRealizacijeMeraTemp)
            var descriptionMera = meraDef.Attribute(Constants.AT_DESC, nLocale).setValue(meraNameTemp)
            var descriptionMera = meraDef.Attribute(Constants.AT_NAME, nLocale).setValue(meraNameTempModified)

        } else {
            //meraDef update attributes 24-29     
            var meraNameTemp = sheets[0].getCell(i, 23-1).getCellValue()
            var meraNameTempModified = meraNameTemp.slice(0,240)
            var opisOcekivanogRezTemp = sheets[0].getCell(i, 24-1).getCellValue()
            var rokIspunjenaAktivnostiTemp = sheets[0].getCell(i, 25-1).getCellValue()
            var odgovoranTemp = sheets[0].getCell(i, 26-1).getCellValue()
            var troskoviAktivnostiTemp = sheets[0].getCell(i, 27-1).getCellValue()
            var opisRealizacijeMeraTemp = sheets[0].getCell(i, 28-1).getCellValue()
            var statusRealizacijeMeraTemp = sheets[0].getCell(i, 29-1).getCellValue()

            var mGroup = riskDef.Group()
            var meraNewObjDef = mGroup.CreateObjDef(Constants.OT_FUNC_INST, meraNameTempModified, nLocale) //kreiraj objDef mere    Initiative		553	ST_INITIATIVE	ARIS default	               
            var setDefSymbol = meraNewObjDef.setDefaultSymbolNum(Constants.ST_INITIATIVE, true)

            var modelTarget = getKpiAllocationDiagram(riskDef)

            var createOccMere = modelTarget.createObjOcc(Constants.ST_INITIATIVE, meraNewObjDef, 100, 100, true, false)

            var g_ntranspoint = new Array()
            var riskSourceOcc = getRiskTarget(riskDef)
            g_ntranspoint[0] = new PublicInterface.Point(createOccMere.X(), createOccMere.Y());
            g_ntranspoint[1] = new PublicInterface.Point(riskSourceOcc.X(), riskSourceOcc.Y());
            var creatCxnOcc = modelTarget.CreateCxnOcc(riskSourceOcc, createOccMere, Constants.CT_HAS_INFL, g_ntranspoint, false);

            var applyTemplateResult = modelTarget.ApplyTemplate()

            var layoutParameters = ArisData.getModelOptions().getLayoutParameters(Constants.MT_KPI_ALLOC_DGM)

            layoutParameters.setArrangeSatellites(2)
            layoutParameters.setSpacingX(300)
            layoutParameters.setSpacingY(100)
            layoutParameters.setQuality(5)

            //currentModel.setAttrOccHandling( ArisData.getModelOptions().getTextAttributeHandling() )
            modelTarget.doLayout()

            var opisOcekivanogRez = meraNewObjDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("74558751-4e17-11ed-7867-00505682218a"), nLocale).setValue(opisOcekivanogRezTemp)
            var rokIspunjenaAktivnosti = meraNewObjDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("88ef6c80-4e17-11ed-7867-00505682218a"), nLocale).setValue(rokIspunjenaAktivnostiTemp)
            var odgovoran = meraNewObjDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("9580bd51-4e17-11ed-7867-00505682218a"), nLocale).setValue(odgovoranTemp)
            var troskoviAktivnosti = meraNewObjDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("a983b9b1-4e17-11ed-7867-00505682218a"), nLocale).setValue(troskoviAktivnostiTemp)
            var opisRealizacijeMera = meraNewObjDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("baa763e0-4e17-11ed-7867-00505682218a"), nLocale).setValue(opisRealizacijeMeraTemp)
            var statusRealizacijeMera = meraNewObjDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("cb0fd461-4e17-11ed-7867-00505682218a"), nLocale).setValue(statusRealizacijeMeraTemp)
            var descriptionMera = meraNewObjDef.Attribute(Constants.AT_DESC, nLocale).setValue(meraNameTemp)
            
            printRow(riskDef.Name(nLocale), "", "", "", meraNewObjDef.Name(nLocale), meraNewObjDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("74558751-4e17-11ed-7867-00505682218a"), nLocale).Type(), opisOcekivanogRezTemp, "NOVA MERA KREIRANA", gExcelObject, rowColorBr)
            printRow(riskDef.Name(nLocale), "", "", "", meraNewObjDef.Name(nLocale), meraNewObjDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("88ef6c80-4e17-11ed-7867-00505682218a"), nLocale).Type(), rokIspunjenaAktivnostiTemp, "NOVA MERA KREIRANA", gExcelObject, rowColorBr)
            printRow(riskDef.Name(nLocale), "", "", "", meraNewObjDef.Name(nLocale), meraNewObjDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("9580bd51-4e17-11ed-7867-00505682218a"), nLocale).Type(), odgovoranTemp, "NOVA MERA KREIRANA", gExcelObject, rowColorBr)
            printRow(riskDef.Name(nLocale), "", "", "", meraNewObjDef.Name(nLocale), meraNewObjDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("a983b9b1-4e17-11ed-7867-00505682218a"), nLocale).Type(), troskoviAktivnostiTemp, "NOVA MERA KREIRANA", gExcelObject, rowColorBr)
            printRow(riskDef.Name(nLocale), "", "", "", meraNewObjDef.Name(nLocale), meraNewObjDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("baa763e0-4e17-11ed-7867-00505682218a"), nLocale).Type(), opisRealizacijeMeraTemp, "NOVA MERA KREIRANA", gExcelObject, rowColorBr)
            printRow(riskDef.Name(nLocale), "", "", "", meraNewObjDef.Name(nLocale), meraNewObjDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("cb0fd461-4e17-11ed-7867-00505682218a"), nLocale).Type(), statusRealizacijeMeraTemp, "NOVA MERA KREIRANA", gExcelObject, rowColorBr)
            
            
        }
        //add attributes on risk object Temp

        var rskDescriptionTemp = sheets[0].getCell(i, 1).getCellValue()
        var faktorRizikaPrilikeTemp = sheets[0].getCell(i, 2).getCellValue()
        var posledicaRizikaTemp = sheets[0].getCell(i, 3).getCellValue()
        
        var orgDeoTemp = sheets[0].getCell(i, 4).getCellValue() //check org

        var imePrezimeVlasnikTemp = sheets[0].getCell(i, 6-1).getCellValue()
        var revidiranjeIzvrsioTemp = sheets[0].getCell(i, 7-1).getCellValue()
        
        var procesTemp = sheets[0].getCell(i, 9-1).getCellValue() //check proces
       
        var pristupProceniTemp = sheets[0].getCell(i, 12-1).getCellValue()
        var koristenaMetodaTemp = sheets[0].getCell(i, 13-1).getCellValue()
        var verovatnocaMetoda1Temp = sheets[0].getCell(i, 14-1).getCellValue()
        var verovatnocaMetoda2Temp = sheets[0].getCell(i, 15-1).getCellValue()
        var verovatnocaMetoda3Temp = sheets[0].getCell(i, 16-1).getCellValue()
        var uticajMetoda1Temp = sheets[0].getCell(i, 17-1).getCellValue()
        var uticajMetoda2Temp = sheets[0].getCell(i, 18-1).getCellValue()
        var uticajMetoda3Temp = sheets[0].getCell(i, 19-1).getCellValue()
        var finUticajTemp = sheets[0].getCell(i, 20-1).getCellValue()
        var nivoTemp = sheets[0].getCell(i, 21-1).getCellValue()
        var odgNaRizikPrilikuTemp = sheets[0].getCell(i, 22-1).getCellValue()
        var rizikRealizacijaTemp = sheets[0].getCell(i, 30-1).getCellValue()
        var opisRealizacijeRizikaTemp = sheets[0].getCell(i, 33-1).getCellValue()
        var pokazateljRizikaTemp = sheets[0].getCell(i, 34-1).getCellValue()
        var uticajRealizacijeRizikaTemp = sheets[0].getCell(i, 35-1).getCellValue()
        var opisKalkulacijeOstvarenjaTemp = sheets[0].getCell(i, 31-1).getCellValue()
        var finUticajDrustvoTemp = sheets[0].getCell(i, 32-1).getCellValue()
        
        var rskDescriptionOld = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("d0dd1fc0-4e0c-11ed-7867-00505682218a"), nLocale).getValue()
        var faktorRizikaPrilikeOld = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("e415a031-4e0c-11ed-7867-00505682218a"), nLocale).getValue()
        var posledicaRizikaOld = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("f4a324e0-4e0c-11ed-7867-00505682218a"), nLocale).getValue()
        var imePrezimeVlasnikOld = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("066e32a1-4e0d-11ed-7867-00505682218a"), nLocale).getValue()
        var revidiranjeIzvrsioOld = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("21e8ce00-4e0d-11ed-7867-00505682218a"), nLocale).getValue()
        var pristupProceniOld = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("4d973550-4e0d-11ed-7867-00505682218a"), nLocale).getValue()
        var koristenaMetodaOld = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("5ee571f1-4e0d-11ed-7867-00505682218a"), nLocale).getValue()
        var verovatnocaMetoda1Old = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("6d660fa1-4e0d-11ed-7867-00505682218a"), nLocale).getValue()
        var verovatnocaMetoda2Old = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("83cc2950-4e0d-11ed-7867-00505682218a"), nLocale).getValue()
        var verovatnocaMetoda3Old = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("b1af9910-4e0d-11ed-7867-00505682218a"), nLocale).getValue()
        var uticajMetoda1Old = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("c148f9c1-4e0d-11ed-7867-00505682218a"), nLocale).getValue()
        var uticajMetoda2Old = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("cdfee991-4e0d-11ed-7867-00505682218a"), nLocale).getValue()
        var uticajMetoda3Old = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("d95aeeb0-4e0d-11ed-7867-00505682218a"), nLocale).getValue()
        var finUticajOld = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("e8d814d1-4e0d-11ed-7867-00505682218a"), nLocale).getValue()
        var nivoOld = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("f78b0d21-4e0d-11ed-7867-00505682218a"), nLocale).getValue()
        var odgNaRizikPrilikuOld = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("08fc6220-4e0e-11ed-7867-00505682218a"), nLocale).getValue()
        var rizikRealizacijaOld = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("1903cd40-4e16-11ed-7867-00505682218a"), nLocale).getValue()
        var opisRealizacijeRizikaOld = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("3f008880-4e16-11ed-7867-00505682218a"), nLocale).getValue()
        var pokazateljRizikaOld = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("4b780fc1-4e16-11ed-7867-00505682218a"), nLocale).getValue()
        var uticajRealizacijeRizikaOld = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("6662a020-4e16-11ed-7867-00505682218a"), nLocale).getValue()
        var opisKalkulacijeOstvarenjaOld = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("261571a1-4e16-11ed-7867-00505682218a"), nLocale).getValue()
        var finUticajDrustvoOld = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("34385ef1-4e16-11ed-7867-00505682218a"), nLocale).getValue()
        
        var orgUnitTemp = riskDef.getConnectedObjs([Constants.OT_ORG_UNIT], Constants.EDGES_INOUT)
        var orgDeoOld = ObjsListNamesToString(orgUnitTemp)
        
        var procesObjFiltered = new Array()
        var kpiAllocationDiagram = getKpiAllocationDiagram(riskDef)
        if(kpiAllocationDiagram != null){
            var procesObj = riskDef.getConnectedObjs([Constants.OT_FUNC], Constants.EDGES_INOUT)
            for (var j = 0; j < procesObj.length; j++) {
                var procesObjFilteredTemp = procesObj[j].OccListInModel(kpiAllocationDiagram)[0].ObjDef()
                procesObjFiltered.push(procesObjFilteredTemp)
            }
        }
        var procesOld = ObjsListNamesToString(procesObjFiltered)
                                                                                                                                                        
        if(rskDescriptionTemp.replaceAll("[\n\r]", "") != rskDescriptionOld.replaceAll("[\n\r]", "")){
                //create LOG                
                printRow(riskDef.Name(nLocale), riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("d0dd1fc0-4e0c-11ed-7867-00505682218a"), nLocale).Type(), rskDescriptionTemp, rskDescriptionOld, "", "", "", "", gExcelObject, rowColorBr)
                rowColorBr++                  
            }
            if(faktorRizikaPrilikeTemp.replaceAll("[\n\r]", "") != faktorRizikaPrilikeOld.replaceAll("[\n\r]", "")){              
                //create LOG                
                printRow(riskDef.Name(nLocale), riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("e415a031-4e0c-11ed-7867-00505682218a"), nLocale).Type(), faktorRizikaPrilikeTemp, faktorRizikaPrilikeOld, "", "", "", "", gExcelObject, rowColorBr)
                rowColorBr++                  
            }
            if(posledicaRizikaTemp.replaceAll("[\n\r]", "") != posledicaRizikaOld.replaceAll("[\n\r]", "")){              
                //create LOG                
                printRow(riskDef.Name(nLocale), riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("f4a324e0-4e0c-11ed-7867-00505682218a"), nLocale).Type(), posledicaRizikaTemp, posledicaRizikaOld, "", "", "", "", gExcelObject, rowColorBr)
                rowColorBr++                  
            }
            if(imePrezimeVlasnikTemp.replaceAll("[\n\r]", "") != imePrezimeVlasnikOld.replaceAll("[\n\r]", "")){              
                //create LOG                
                printRow(riskDef.Name(nLocale), riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("066e32a1-4e0d-11ed-7867-00505682218a"), nLocale).Type(), imePrezimeVlasnikTemp, imePrezimeVlasnikOld, "", "", "", "", gExcelObject, rowColorBr)
                rowColorBr++                  
            }
            if(revidiranjeIzvrsioTemp.replaceAll("[\n\r]", "") != revidiranjeIzvrsioOld.replaceAll("[\n\r]", "")){              
                //create LOG                
                printRow(riskDef.Name(nLocale), riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("21e8ce00-4e0d-11ed-7867-00505682218a"), nLocale).Type(), revidiranjeIzvrsioTemp, revidiranjeIzvrsioOld, "", "", "", "", gExcelObject, rowColorBr)
                rowColorBr++                  
            }
            if(pristupProceniTemp.replaceAll("[\n\r]", "") != pristupProceniOld.replaceAll("[\n\r]", "")){              
                //create LOG                
                printRow(riskDef.Name(nLocale), riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("4d973550-4e0d-11ed-7867-00505682218a"), nLocale).Type(), pristupProceniTemp, pristupProceniOld, "", "", "", "", gExcelObject, rowColorBr)
                rowColorBr++                  
            }
            if(koristenaMetodaTemp.replaceAll("[\n\r]", "") != koristenaMetodaOld.replaceAll("[\n\r]", "")){              
                //create LOG                
                printRow(riskDef.Name(nLocale), riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("5ee571f1-4e0d-11ed-7867-00505682218a"), nLocale).Type(), koristenaMetodaTemp, koristenaMetodaOld, "", "", "", "", gExcelObject, rowColorBr)
                rowColorBr++                  
            }
            if(verovatnocaMetoda1Temp.replaceAll("[\n\r]", "") != verovatnocaMetoda1Old.replaceAll("[\n\r]", "")){              
                //create LOG                
                printRow(riskDef.Name(nLocale), riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("6d660fa1-4e0d-11ed-7867-00505682218a"), nLocale).Type(), verovatnocaMetoda1Temp, verovatnocaMetoda1Old, "", "", "", "", gExcelObject, rowColorBr)
                rowColorBr++                  
            }
            if(verovatnocaMetoda2Temp.replaceAll("[\n\r]", "") != verovatnocaMetoda2Old.replaceAll("[\n\r]", "")){              
                //create LOG                
                printRow(riskDef.Name(nLocale), riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("83cc2950-4e0d-11ed-7867-00505682218a"), nLocale).Type(), verovatnocaMetoda2Temp, verovatnocaMetoda2Old, "", "", "", "", gExcelObject, rowColorBr)
                rowColorBr++                  
            }
            if(verovatnocaMetoda3Temp.replaceAll("[\n\r]", "") != verovatnocaMetoda3Old.replaceAll("[\n\r]", "")){              
                //create LOG                
                printRow(riskDef.Name(nLocale), riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("b1af9910-4e0d-11ed-7867-00505682218a"), nLocale).Type(), verovatnocaMetoda3Temp, verovatnocaMetoda3Old, "", "", "", "", gExcelObject, rowColorBr)
                rowColorBr++                  
            }
            if(uticajMetoda1Temp.replaceAll("[\n\r]", "") != uticajMetoda1Old.replaceAll("[\n\r]", "")){              
                //create LOG                
                printRow(riskDef.Name(nLocale), riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("c148f9c1-4e0d-11ed-7867-00505682218a"), nLocale).Type(), uticajMetoda1Temp, uticajMetoda1Old, "", "", "", "", gExcelObject, rowColorBr)
                rowColorBr++                  
            }
            if(uticajMetoda2Temp.replaceAll("[\n\r]", "") != uticajMetoda2Old.replaceAll("[\n\r]", "")){              
                //create LOG                
                printRow(riskDef.Name(nLocale), riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("cdfee991-4e0d-11ed-7867-00505682218a"), nLocale).Type(), uticajMetoda2Temp, uticajMetoda2Old, "", "", "", "", gExcelObject, rowColorBr)
                rowColorBr++                  
            }
            if(uticajMetoda3Temp.replaceAll("[\n\r]", "") != uticajMetoda3Old.replaceAll("[\n\r]", "")){              
                //create LOG                
                printRow(riskDef.Name(nLocale), riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("d95aeeb0-4e0d-11ed-7867-00505682218a"), nLocale).Type(), uticajMetoda3Temp, uticajMetoda3Old, "", "", "", "", gExcelObject, rowColorBr)
                rowColorBr++                  
            }
            if(finUticajTemp.replaceAll("[\n\r]", "") != finUticajOld.replaceAll("[\n\r]", "")){              
                //create LOG                
                printRow(riskDef.Name(nLocale), riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("e8d814d1-4e0d-11ed-7867-00505682218a"), nLocale).Type(), finUticajTemp, finUticajOld, "", "", "", "", gExcelObject, rowColorBr)
                rowColorBr++                  
            }
            if(nivoTemp.replaceAll("[\n\r]", "") != nivoOld.replaceAll("[\n\r]", "")){              
                //create LOG                
                printRow(riskDef.Name(nLocale), riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("f78b0d21-4e0d-11ed-7867-00505682218a"), nLocale).Type(), nivoTemp, nivoOld, "", "", "", "", gExcelObject, rowColorBr)
                rowColorBr++                  
            }
            if(odgNaRizikPrilikuTemp.replaceAll("[\n\r]", "") != odgNaRizikPrilikuOld.replaceAll("[\n\r]", "")){              
                //create LOG                
                printRow(riskDef.Name(nLocale), riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("08fc6220-4e0e-11ed-7867-00505682218a"), nLocale).Type(), odgNaRizikPrilikuTemp, odgNaRizikPrilikuOld, "", "", "", "", gExcelObject, rowColorBr)
                rowColorBr++                  
            }
            if(rizikRealizacijaTemp.replaceAll("[\n\r]", "") != rizikRealizacijaOld.replaceAll("[\n\r]", "")){              
                //create LOG                
                printRow(riskDef.Name(nLocale), riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("1903cd40-4e16-11ed-7867-00505682218a"), nLocale).Type(), rizikRealizacijaTemp, rizikRealizacijaOld, "", "", "", "", gExcelObject, rowColorBr)
                rowColorBr++                  
            }
            if(opisRealizacijeRizikaTemp.replaceAll("[\n\r]", "") != opisRealizacijeRizikaOld.replaceAll("[\n\r]", "")){              
                //create LOG                
                printRow(riskDef.Name(nLocale), riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("3f008880-4e16-11ed-7867-00505682218a"), nLocale).Type(), opisRealizacijeRizikaTemp, opisRealizacijeRizikaOld, "", "", "", "", gExcelObject, rowColorBr)
                rowColorBr++                  
            }
            if(pokazateljRizikaTemp.replaceAll("[\n\r]", "") != pokazateljRizikaOld.replaceAll("[\n\r]", "")){              
                //create LOG                
                printRow(riskDef.Name(nLocale), riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("4b780fc1-4e16-11ed-7867-00505682218a"), nLocale).Type(), pokazateljRizikaTemp, pokazateljRizikaOld, "", "", "", "", gExcelObject, rowColorBr)
                rowColorBr++                  
            }
            if(uticajRealizacijeRizikaTemp.replaceAll("[\n\r]", "") != uticajRealizacijeRizikaOld.replaceAll("[\n\r]", "")){              
                //create LOG                
                printRow(riskDef.Name(nLocale), riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("6662a020-4e16-11ed-7867-00505682218a"), nLocale).Type(), uticajRealizacijeRizikaTemp, uticajRealizacijeRizikaOld, "", "", "", "", gExcelObject, rowColorBr)
                rowColorBr++                  
            }
            if(opisKalkulacijeOstvarenjaTemp.replaceAll("[\n\r]", "") != opisKalkulacijeOstvarenjaOld.replaceAll("[\n\r]", "")){              
                //create LOG                
                printRow(riskDef.Name(nLocale), riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("261571a1-4e16-11ed-7867-00505682218a"), nLocale).Type(), opisKalkulacijeOstvarenjaTemp, opisKalkulacijeOstvarenjaOld, "", "", "", "", gExcelObject, rowColorBr)
                rowColorBr++                  
            }
            if(finUticajDrustvoTemp.replaceAll("[\n\r]", "") != finUticajDrustvoOld.replaceAll("[\n\r]", "")){              
                //create LOG                
                printRow(riskDef.Name(nLocale), riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("34385ef1-4e16-11ed-7867-00505682218a"), nLocale).Type(), finUticajDrustvoTemp, finUticajDrustvoOld, "", "", "", "", gExcelObject, rowColorBr)
                rowColorBr++                  
            }
             if(orgDeoTemp.replaceAll("[\n\r\u2022]", "") != orgDeoOld.replace(/[\n\r\u2022]/g, "")){              
                //create LOG                
                printRow(riskDef.Name(nLocale), "Objekti organizacije", orgDeoTemp, orgDeoOld, "", "", "", "", gExcelObject, rowColorBr)
                rowColorBr++                  
            }
            if(procesTemp.replaceAll("[\n\r\u2022]", "") != procesOld.replace(/[\n\r\u2022]/g, "")){              
                //create LOG                
                printRow(riskDef.Name(nLocale), "Objekti procesa", procesTemp, procesOld, "", "", "", "", gExcelObject, rowColorBr)
                rowColorBr++                  
            }

        var rskDescription = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("d0dd1fc0-4e0c-11ed-7867-00505682218a"), nLocale).setValue(rskDescriptionTemp)
        var faktorRizikaPrilike = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("e415a031-4e0c-11ed-7867-00505682218a"), nLocale).setValue(faktorRizikaPrilikeTemp)
        var posledicaRizika = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("f4a324e0-4e0c-11ed-7867-00505682218a"), nLocale).setValue(posledicaRizikaTemp)
        var imePrezimeVlasnik = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("066e32a1-4e0d-11ed-7867-00505682218a"), nLocale).setValue(imePrezimeVlasnikTemp)
        var revidiranjeIzvrsio = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("21e8ce00-4e0d-11ed-7867-00505682218a"), nLocale).setValue(revidiranjeIzvrsioTemp)
        var pristupProceni = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("4d973550-4e0d-11ed-7867-00505682218a"), nLocale).setValue(pristupProceniTemp)
        var koristenaMetoda = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("5ee571f1-4e0d-11ed-7867-00505682218a"), nLocale).setValue(koristenaMetodaTemp)
        var verovatnocaMetoda1 = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("6d660fa1-4e0d-11ed-7867-00505682218a"), nLocale).setValue(verovatnocaMetoda1Temp)
        var verovatnocaMetoda2 = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("83cc2950-4e0d-11ed-7867-00505682218a"), nLocale).setValue(verovatnocaMetoda2Temp)
        var verovatnocaMetoda3 = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("b1af9910-4e0d-11ed-7867-00505682218a"), nLocale).setValue(verovatnocaMetoda3Temp)
        var uticajMetoda1 = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("c148f9c1-4e0d-11ed-7867-00505682218a"), nLocale).setValue(uticajMetoda1Temp)
        var uticajMetoda2 = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("cdfee991-4e0d-11ed-7867-00505682218a"), nLocale).setValue(uticajMetoda2Temp)
        var uticajMetoda3 = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("d95aeeb0-4e0d-11ed-7867-00505682218a"), nLocale).setValue(uticajMetoda3Temp)
        var finUticaj = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("e8d814d1-4e0d-11ed-7867-00505682218a"), nLocale).setValue(finUticajTemp)
        var nivo = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("f78b0d21-4e0d-11ed-7867-00505682218a"), nLocale).setValue(nivoTemp)
        var odgNaRizikPriliku = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("08fc6220-4e0e-11ed-7867-00505682218a"), nLocale).setValue(odgNaRizikPrilikuTemp)
        var rizikRealizacija = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("1903cd40-4e16-11ed-7867-00505682218a"), nLocale).setValue(rizikRealizacijaTemp)
        var opisRealizacijeRizika = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("3f008880-4e16-11ed-7867-00505682218a"), nLocale).setValue(opisRealizacijeRizikaTemp)
        var pokazateljRizika = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("4b780fc1-4e16-11ed-7867-00505682218a"), nLocale).setValue(pokazateljRizikaTemp)
        var uticajRealizacijeRizika = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("6662a020-4e16-11ed-7867-00505682218a"), nLocale).setValue(uticajRealizacijeRizikaTemp)
        var opisKalkulacijeOstvarenja = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("261571a1-4e16-11ed-7867-00505682218a"), nLocale).setValue(opisKalkulacijeOstvarenjaTemp)
        var finUticajDrustvo = riskDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("34385ef1-4e16-11ed-7867-00505682218a"), nLocale).setValue(finUticajDrustvoTemp)

    }
}
gExcelObject.write();

function getKpiAllocationDiagram(riskDef) {
    var occsTemp = riskDef.OccList()
    var modelKpi = null
    for (var i = 0; i < occsTemp.length; i++) {
        var tempModel = occsTemp[i].Model()

        if (tempModel.TypeNum() == Constants.MT_KPI_ALLOC_DGM) {
            modelKpi = tempModel
        }
    }
    return modelKpi
}

function getRiskTarget(riskDef) {
    var occsTemp = riskDef.OccList()
    var riskTarget = null

    for (var i = 0; i < occsTemp.length; i++) {
        var tempModel = occsTemp[i].Model()

        if (tempModel.TypeNum() == Constants.MT_KPI_ALLOC_DGM) {
            riskTarget = occsTemp[i]
        }
    }
    return riskTarget
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

function printRowHeaders(pExcelWorkBook) {
    var sDate = java.text.SimpleDateFormat("dd-MM-yyyy HH.mm").format(new java.util.Date());
    var sSheetName = gsSheetName + sDate;
    var lMainSheet = pExcelWorkBook.createSheet(sSheetName);
    var lHeaderRow = lMainSheet.createRow(giRowCounter);
    // set column width
    /*
    // Part of
    lMainSheet.setColumnWidth(0, giStandardColumnWidth);
    // Activity name
    lMainSheet.setColumnWidth(1, giStandardColumnWidth);
    // Activity Description
    lMainSheet.setColumnWidth(2, giStandardColumnWidth);
    // Input
    lMainSheet.setColumnWidth(3, giStandardColumnWidth);
    // Output
    lMainSheet.setColumnWidth(4, giStandardColumnWidth);
    */

    lMainSheet.setDefaultColumnWidth(30);

    /* Freeze top row alone */
    lMainSheet.createFreezePane(0, 1);

    // set style
    var lHeaderFont = pExcelWorkBook.createFont();
    lHeaderFont.setBoldweight(Constants.XL_FONT_WEIGHT_BOLD);
    var headLiBorderColor = Constants.C_GRAY;

    headerStyle = pExcelWorkBook.createCellStyle(lHeaderFont, 1, 1, 1, 1, headLiBorderColor, headLiBorderColor, headLiBorderColor, headLiBorderColor, Constants.ALIGN_CENTER, Constants.VERTICAL_CENTER, Constants.C_TRANSPARENT, Constants.C_LIGHT_BLUE, Constants.SOLID_FOREGROUND);
    headerStyle.setWrapText(true)


    var lHeaderFontGuid = pExcelWorkBook.createFont();
    lHeaderFontGuid.setBoldweight(Constants.XL_FONT_WEIGHT_BOLD);
    lHeaderFontGuid.setColor(10);
    var headLiBorderColorGuid = Constants.C_GRAY;

    headerStyleGuid = pExcelWorkBook.createCellStyle(lHeaderFontGuid, 1, 1, 1, 1, headLiBorderColorGuid, headLiBorderColorGuid, headLiBorderColorGuid, headLiBorderColorGuid, Constants.ALIGN_CENTER, Constants.VERTICAL_CENTER, Constants.C_TRANSPARENT, Constants.C_BLACK, Constants.SOLID_FOREGROUND);
    headerStyleGuid.setWrapText(true)


    var lCellL1 = lHeaderRow.createCell(0);
    lCellL1.setCellValue("Naziv rizika");
    lCellL1.setCellStyle(headerStyle);

    //var lCellStyle = lCellL1.getCellStyle();
    //lCellStyle.setFont(lHeaderFont);

    var lCellL2 = lHeaderRow.createCell(1);
    lCellL2.setCellValue("Atribut rizika / Objekt");
    lCellL2.setCellStyle(headerStyle);

    var lCellL3 = lHeaderRow.createCell(2);
    lCellL3.setCellValue("Nova vrednost");
    lCellL3.setCellStyle(headerStyle);

    var lCellL4 = lHeaderRow.createCell(3);
    lCellL4.setCellValue("Stara vrednost");
    lCellL4.setCellStyle(headerStyle);

    var lCell5 = lHeaderRow.createCell(4);
    lCell5.setCellValue("Naziv mere");
    lCell5.setCellStyle(headerStyle);

    var lCell6 = lHeaderRow.createCell(5);
    lCell6.setCellValue("Atribut mere");
    lCell6.setCellStyle(headerStyle);

    var lCell7 = lHeaderRow.createCell(6);
    lCell7.setCellValue("Nova vrednost");
    lCell7.setCellStyle(headerStyle);
    
    var lCell8 = lHeaderRow.createCell(7);
    lCell8.setCellValue("Stara vrednost");
    lCell8.setCellStyle(headerStyle);

    // set new style
    var lDataFont = pExcelWorkBook.createFont();
    lDataFont.setBoldweight(Constants.XL_FONT_WEIGHT_NORMAL);

    var liBorderColor = Constants.C_GRAY;
    lDataRowStyle = pExcelWorkBook.createCellStyle(lDataFont, 1, 1, 1, 1, liBorderColor, liBorderColor, liBorderColor, liBorderColor, Constants.ALIGN_LEFT, Constants.VERTICAL_TOP, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.SOLID_FOREGROUND);
    lDataRowStyle.setWrapText(true);
    // set style
    /* var lCell13 = lHeaderRow.createCell(12);
    lCell13.setCellStyle(lDataRowStyle);
    lCell13.setCellValue(" "); */

    giRowCounter++
}

function printRow(riskName, riskAtt, riskAttNew, riskAttOld, meraName, meraAtt, meraAttNew, meraAttOld, pExcelWorkBook, rowColorBr) {
    // try {


    //shade every other row, based on rowColorBr counter
    var crFont = pExcelWorkBook.createFont();
    crFont.setBoldweight(Constants.XL_FONT_WEIGHT_NORMAL);
    var crLiBorderColor = Constants.C_GRAY;

    crDataRowStyle = pExcelWorkBook.createCellStyle(crFont, 1, 1, 1, 1, crLiBorderColor, crLiBorderColor, crLiBorderColor, crLiBorderColor, Constants.ALIGN_LEFT, Constants.VERTICAL_TOP, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.SOLID_FOREGROUND);

    /* if (rowColorBr % 2 == 1) {
           crDataRowStyle = pExcelWorkBook.createCellStyle(crFont, 1, 1, 1, 1, crLiBorderColor, crLiBorderColor, crLiBorderColor, crLiBorderColor, Constants.ALIGN_LEFT, Constants.VERTICAL_TOP, Constants.C_TRANSPARENT, Constants.C_GREY_80_PERCENT, Constants.SOLID_FOREGROUND);
       } else {
           crDataRowStyle = pExcelWorkBook.createCellStyle(crFont, 1, 1, 1, 1, crLiBorderColor, crLiBorderColor, crLiBorderColor, crLiBorderColor, Constants.ALIGN_LEFT, Constants.VERTICAL_TOP, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.SOLID_FOREGROUND);
       } 
    */

    crDataRowStyle.setWrapText(true);

    //GUID COLOR
    var lFontGuid = pExcelWorkBook.createFont();
    //lFontGuid.setBoldweight(Constants.XL_FONT_WEIGHT_BOLD);
    lFontGuid.setColor(10);
    var iBorderColorGuid = Constants.C_GRAY;

    iStyleGuid = pExcelWorkBook.createCellStyle(lFontGuid, 1, 1, 1, 1, iBorderColorGuid, iBorderColorGuid, iBorderColorGuid, iBorderColorGuid, Constants.ALIGN_CENTER, Constants.VERTICAL_CENTER, Constants.C_TRANSPARENT, Constants.C_BLACK, Constants.SOLID_FOREGROUND);
    iStyleGuid.setWrapText(true)

    var lMainSheet = pExcelWorkBook.getSheetAt(0);

    // second row data
    var lModelRow = lMainSheet.createRow(giRowCounter);

    var lCell = lModelRow.createCell(0);
    lCell.setCellValue(riskName);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(1);
    lCell.setCellValue(riskAtt);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(2);
    lCell.setCellValue(riskAttNew);
    lCell.setCellStyle(crDataRowStyle);

    // Input Name
    var lCell = lModelRow.createCell(3);
    lCell.setCellValue(riskAttOld);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(4);
    lCell.setCellValue(meraName);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(5);
    lCell.setCellValue(meraAtt);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(6);
    lCell.setCellValue(meraAttNew);
    lCell.setCellStyle(crDataRowStyle);
    
    var lCell = lModelRow.createCell(7);
    lCell.setCellValue(meraAttOld);
    lCell.setCellStyle(crDataRowStyle);

    // up counter
    giRowCounter++;
    //} catch (ex) {
    //   Dialogs.MsgBox("Error : " + ex);
    // }
}

function ObjsListNamesToString(paObjList) {
    var psResult = "";
    var liCounter = 0;
    var nrOfElements = paObjList.length;
    for each(loObj in paObjList) {
        if (liCounter + 1 == nrOfElements) {
            psResult += gsBulletSign + " " + loObj.Name(nLocale);
        } else {
            psResult += gsBulletSign + " " + loObj.Name(nLocale) + "\n";
        }
    }
    return psResult;
}

function getKpiAllocationDiagram(riskDef) {
    var occsTemp = riskDef.OccList()

    var modelKpi = null
    for (var i = 0; i < occsTemp.length; i++) {
        var tempModel = occsTemp[i].Model()
        if (tempModel.TypeNum() == Constants.MT_KPI_ALLOC_DGM) {
            modelKpi = tempModel
        }
    }
    return modelKpi
}