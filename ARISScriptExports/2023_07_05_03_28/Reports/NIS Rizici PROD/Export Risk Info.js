// static variables
var giLoc = Context.getSelectedLanguage();
var gExcelObject = null;
var giStandardColumnWidth = 15000;
var giRowCounter = 0;
var lDataRowStyle = null;
var gCurrentFilter;
var gsBulletSign = "\u2022";
var gsSheetName = "Printed at ";

function main() {
    gExcelObject = Context.createExcelWorkbook(Context.getSelectedFile());

    var selectedModels = ArisData.getSelectedModels();
    var selectedObjects = ArisData.getSelectedObjDefs();

    if (selectedObjects.length > 0) {
        // save active metod filter
        gCurrentFilter = ArisData.ActiveFilter();
        printRowHeaders(gExcelObject);
        // loop through selected models and output results
        for (var i = 0; i < selectedObjects.length; i++) {
            collectAndPrintData(selectedObjects[i], gExcelObject);
        }
    } else if (selectedModels.length > 0) {

        gCurrentFilter = ArisData.ActiveFilter();
        printRowHeaders(gExcelObject);

        var selectedObjects = selectedModels[0].ObjDefListFilter(Constants.OT_RISK)

        // loop through selected objects and output results
        for (var i = 0; i < selectedObjects.length; i++) {
            collectAndPrintData(selectedObjects[i], gExcelObject);
        }

    }
    // write report
    gExcelObject.write();
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
    lCellL2.setCellValue("Opis rizika/prilike");
    lCellL2.setCellStyle(headerStyle);

    var lCellL3 = lHeaderRow.createCell(2);
    lCellL3.setCellValue("Faktori rizika/prilike");
    lCellL3.setCellStyle(headerStyle);

    var lCellL4 = lHeaderRow.createCell(3);
    lCellL4.setCellValue("Posledica");
    lCellL4.setCellStyle(headerStyle);

    var lCell5 = lHeaderRow.createCell(4);
    lCell5.setCellValue("Orgаnizacioni deo");
    lCell5.setCellStyle(headerStyle);

    //var lCell6 = lHeaderRow.createCell(5);
    //lCell6.setCellValue("Pozicija Vlasnik rizika/prilike");
    //lCell6.setCellStyle(headerStyle);

    var lCell7 = lHeaderRow.createCell(6-1);
    lCell7.setCellValue("Ime I prezime Vlasnik rizika/prilike");
    lCell7.setCellStyle(headerStyle);

    var lCell8 = lHeaderRow.createCell(7-1);
    lCell8.setCellValue("Revidiranje izvršio");
    lCell8.setCellStyle(headerStyle);

    var lCell9 = lHeaderRow.createCell(8-1);
    lCell9.setCellValue("Kategorija biznis procesa");
    lCell9.setCellStyle(headerStyle);

    var lCell10 = lHeaderRow.createCell(9-1);
    lCell10.setCellValue("Proces");
    lCell10.setCellStyle(headerStyle);

    var lCell11 = lHeaderRow.createCell(10-1);
    lCell11.setCellValue("Grupa rizika");
    lCell11.setCellStyle(headerStyle);

    var lCell12 = lHeaderRow.createCell(11-1);
    lCell12.setCellValue("Podgrupa rizika");
    lCell12.setCellStyle(headerStyle);

    var lCell13 = lHeaderRow.createCell(12-1);
    lCell13.setCellValue("Pristup proceni");
    lCell13.setCellStyle(headerStyle);

    var lCell14 = lHeaderRow.createCell(13-1);
    lCell14.setCellValue("Korišćena metoda");
    lCell14.setCellStyle(headerStyle);

    var lCell15 = lHeaderRow.createCell(14-1);
    lCell15.setCellValue("Verovatnoća \n(stepen neizvesnosti da li će se rizik desiti) \nMetoda 1 ЕМВ ocena");
    lCell15.setCellStyle(headerStyle);

    var lCell16 = lHeaderRow.createCell(15-1);
    lCell16.setCellValue("Verovatnoća \n(stepen neizvesnosti da li će se rizik desiti) \nMetoda 2 Kvantitativna ocena");
    lCell16.setCellStyle(headerStyle);

    var lCell17 = lHeaderRow.createCell(16-1);
    lCell17.setCellValue("Verovatnoća \n(stepen neizvesnosti da li će se rizik desiti) \nMetoda 3 Kvalitativna ocena");
    lCell17.setCellStyle(headerStyle);

    var lCell18 = lHeaderRow.createCell(17-1);
    lCell18.setCellValue("Uticaj \n(stepen uticaja rizika na poslovanje Društva) \nMetoda 1 ЕМВ ocena");
    lCell18.setCellStyle(headerStyle);

    var lCell19 = lHeaderRow.createCell(18-1);
    lCell19.setCellValue("Uticaj \n(stepen uticaja rizika na poslovanje Društva) \nMetoda 2 Kvantitativna ocena");
    lCell19.setCellStyle(headerStyle);

    var lCell20 = lHeaderRow.createCell(19-1);
    lCell20.setCellValue("Uticaj \n(stepen uticaja rizika na poslovanje Društva) \nMetoda 3 Kvalitativna ocena");
    lCell20.setCellStyle(headerStyle);

    var lCell21 = lHeaderRow.createCell(20-1);
    lCell21.setCellValue("Finansijski uticaj (mln. din)");
    lCell21.setCellStyle(headerStyle);

    var lCell22 = lHeaderRow.createCell(21-1);
    lCell22.setCellValue("Nivo");
    lCell22.setCellStyle(headerStyle);

    var lCell23 = lHeaderRow.createCell(22-1);
    lCell23.setCellValue("Odgovor na rizik/priliku");
    lCell23.setCellStyle(headerStyle);

    var lCell24 = lHeaderRow.createCell(23-1);
    lCell24.setCellValue("Mere za upravljanje rizicima");
    lCell24.setCellStyle(headerStyle);

    var lCell25 = lHeaderRow.createCell(24-1);
    lCell25.setCellValue("Opis očekivanog rezultata");
    lCell25.setCellStyle(headerStyle);

    var lCell26 = lHeaderRow.createCell(25-1);
    lCell26.setCellValue("Rok ispunjenja aktivnosti");
    lCell26.setCellStyle(headerStyle);

    var lCell27 = lHeaderRow.createCell(26-1);
    lCell27.setCellValue("Odgovoran");
    lCell27.setCellStyle(headerStyle);

    var lCell28 = lHeaderRow.createCell(27-1);
    lCell28.setCellValue("Troškovi za aktivnosti");
    lCell28.setCellStyle(headerStyle);

    var lCell29 = lHeaderRow.createCell(28-1);
    lCell29.setCellValue("Оpis realizacije mera");
    lCell29.setCellStyle(headerStyle);

    var lCell30 = lHeaderRow.createCell(29-1);
    lCell30.setCellValue("Status realizacije mera");
    lCell30.setCellStyle(headerStyle);

    var lCell31 = lHeaderRow.createCell(30-1);
    lCell31.setCellValue("Da li se rizik/ prilika relizovao?");
    lCell31.setCellStyle(headerStyle);

    var lCell32 = lHeaderRow.createCell(31-1);
    lCell32.setCellValue("1. Kvantitativna metoda \n(prioritet prilikom ocene realiizacije rizika/ prilike) \nOpis kalkulacije ostvarenja rizika/prilike \n(opisati način na koji je iznos ostvarenog rizika/prilike izračunat)");
    lCell32.setCellStyle(headerStyle);

    var lCell33 = lHeaderRow.createCell(32-1);
    lCell33.setCellValue("1. Kvantitativna metoda \n(prioritet prilikom ocene realiizacije rizika/ prilike) \nFinansijski uticaj na Društvo (mil. din.)");
    lCell33.setCellStyle(headerStyle);

    var lCell34 = lHeaderRow.createCell(33-1);
    lCell34.setCellValue("2. Kvalititativna metoda \n(ukoliko realizaciju rizika/prilike nije moguće oceniti kvantitativnom metodom) \nOpis realizacije rizika/prilike");
    lCell34.setCellStyle(headerStyle);

    var lCell35 = lHeaderRow.createCell(34-1);
    lCell35.setCellValue("2. Kvalititativna metoda \n(ukoliko realizaciju rizika/prilike nije moguće oceniti kvantitativnom metodom) \nPokazatelj rizika/prilike");
    lCell35.setCellStyle(headerStyle);

    var lCell36 = lHeaderRow.createCell(35-1);
    lCell36.setCellValue("Uticaj realizacije rizika/prilike na poslovne pokazatelje Društva");
    lCell36.setCellStyle(headerStyle);

    var lCell37 = lHeaderRow.createCell(36-1);
    lCell37.setCellValue("RIZIK GUID");
    lCell37.setCellStyle(headerStyleGuid);

    var lCell38 = lHeaderRow.createCell(37-1);
    lCell38.setCellValue("MERA GUID");
    lCell38.setCellStyle(headerStyleGuid);

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

function collectAndPrintData(pObject, pExcelWorkBook) {
    //try {

    var rowColorBr = 1;

    // activity counter, used for merge
    var funcBr = 0;
    var defaultValue = ""

    var rskGuid = pObject.GUID()
    var rskObjectName = pObject.Name(giLoc);
    var rskDescription = getItemAttributeValue(pObject, ArisData.ActiveFilter().UserDefinedAttributeTypeNum("d0dd1fc0-4e0c-11ed-7867-00505682218a"), defaultValue)
    var faktorRizikaPrilike = getItemAttributeValue(pObject, ArisData.ActiveFilter().UserDefinedAttributeTypeNum("e415a031-4e0c-11ed-7867-00505682218a"), defaultValue)
    var posledicaRizika = getItemAttributeValue(pObject, ArisData.ActiveFilter().UserDefinedAttributeTypeNum("f4a324e0-4e0c-11ed-7867-00505682218a"), defaultValue)

    var orgUnitTemp = pObject.getConnectedObjs([Constants.OT_ORG_UNIT], Constants.EDGES_INOUT)
    var orgDeo = ObjsListNamesToString(orgUnitTemp)

    var positionTemp = pObject.getConnectedObjs([Constants.OT_POS], Constants.EDGES_INOUT)
    var pozicijaVlasnik = ObjsListNamesToString(positionTemp) //obj pozicija

    var imePrezimeVlasnik = getItemAttributeValue(pObject, ArisData.ActiveFilter().UserDefinedAttributeTypeNum("066e32a1-4e0d-11ed-7867-00505682218a"), defaultValue)
    var revidiranjeIzvrsio = getItemAttributeValue(pObject, ArisData.ActiveFilter().UserDefinedAttributeTypeNum("21e8ce00-4e0d-11ed-7867-00505682218a"), defaultValue)

    var procesObjFiltered = new Array()
    var kpiAllocationDiagram = getKpiAllocationDiagram(pObject)
    if(kpiAllocationDiagram != null){
        var procesObj = pObject.getConnectedObjs([Constants.OT_FUNC], Constants.EDGES_INOUT)
        for (var i = 0; i < procesObj.length; i++) {
            try{
                var procesObjFilteredTemp = procesObj[i].OccListInModel(kpiAllocationDiagram)[0].ObjDef()
                procesObjFiltered.push(procesObjFilteredTemp)
            }
            catch(e){
                //error          
            }
        }
    }
    var proces = ObjsListNamesToString(procesObjFiltered)
   
    var katProcesObj = new Array()   
    if(procesObjFiltered.length > 0){
        var katProcesObj = procesObjFiltered[0].getConnectedObjs([Constants.OT_FUNC], Constants.EDGES_IN, [Constants.CT_IS_PRCS_ORNT_SUPER])
    }
    var katBiznisProcesa = ObjsListNamesToString(katProcesObj) //obj procesa vise razine

    var podgrupaRizikaObj = pObject.getConnectedObjs([Constants.OT_RISK_CATEGORY], Constants.EDGES_INOUT)
    var podgrupaRizika = ObjsListNamesToString(podgrupaRizikaObj)
    
    var grupaRizikaObj = new Array()
    if(podgrupaRizikaObj.length > 0){
        var grupaRizikaObj = podgrupaRizikaObj[0].getConnectedObjs([Constants.OT_RISK_CATEGORY], Constants.EDGES_INOUT)
    }
    var grupaRizika = ObjsListNamesToString(grupaRizikaObj)


    var pristupProceni = getItemAttributeValue(pObject, ArisData.ActiveFilter().UserDefinedAttributeTypeNum("4d973550-4e0d-11ed-7867-00505682218a"), defaultValue)
    var koristenaMetoda = getItemAttributeValue(pObject, ArisData.ActiveFilter().UserDefinedAttributeTypeNum("5ee571f1-4e0d-11ed-7867-00505682218a"), defaultValue)

    var verovatnocaMetoda1 = getItemAttributeValue(pObject, ArisData.ActiveFilter().UserDefinedAttributeTypeNum("6d660fa1-4e0d-11ed-7867-00505682218a"), defaultValue)
    var verovatnocaMetoda2 = getItemAttributeValue(pObject, ArisData.ActiveFilter().UserDefinedAttributeTypeNum("83cc2950-4e0d-11ed-7867-00505682218a"), defaultValue)
    var verovatnocaMetoda3 = getItemAttributeValue(pObject, ArisData.ActiveFilter().UserDefinedAttributeTypeNum("b1af9910-4e0d-11ed-7867-00505682218a"), defaultValue)

    var uticajMetoda1 = getItemAttributeValue(pObject, ArisData.ActiveFilter().UserDefinedAttributeTypeNum("c148f9c1-4e0d-11ed-7867-00505682218a"), defaultValue)
    var uticajMetoda2 = getItemAttributeValue(pObject, ArisData.ActiveFilter().UserDefinedAttributeTypeNum("cdfee991-4e0d-11ed-7867-00505682218a"), defaultValue)
    var uticajMetoda3 = getItemAttributeValue(pObject, ArisData.ActiveFilter().UserDefinedAttributeTypeNum("d95aeeb0-4e0d-11ed-7867-00505682218a"), defaultValue)

    var finUticaj = getItemAttributeValue(pObject, ArisData.ActiveFilter().UserDefinedAttributeTypeNum("e8d814d1-4e0d-11ed-7867-00505682218a"), defaultValue)
    
    if(koristenaMetoda.toString() == "3"){
        var nivo = ""   
    }
    else{
        var nivo = getItemAttributeValue(pObject, ArisData.ActiveFilter().UserDefinedAttributeTypeNum("f78b0d21-4e0d-11ed-7867-00505682218a"), defaultValue)
    }
    
    var odgNaRizikPriliku = getItemAttributeValue(pObject, ArisData.ActiveFilter().UserDefinedAttributeTypeNum("08fc6220-4e0e-11ed-7867-00505682218a"), defaultValue)

    var rizikRealizacija = getItemAttributeValue(pObject, ArisData.ActiveFilter().UserDefinedAttributeTypeNum("1903cd40-4e16-11ed-7867-00505682218a"), defaultValue)
    var opisRealizacijeRizika = getItemAttributeValue(pObject, ArisData.ActiveFilter().UserDefinedAttributeTypeNum("3f008880-4e16-11ed-7867-00505682218a"), defaultValue)
    var pokazateljRizika = getItemAttributeValue(pObject, ArisData.ActiveFilter().UserDefinedAttributeTypeNum("4b780fc1-4e16-11ed-7867-00505682218a"), defaultValue)
    var uticajRealizacijeRizika = getItemAttributeValue(pObject, ArisData.ActiveFilter().UserDefinedAttributeTypeNum("6662a020-4e16-11ed-7867-00505682218a"), defaultValue)

    var opisKalkulacijeOstvarenja = getItemAttributeValue(pObject, ArisData.ActiveFilter().UserDefinedAttributeTypeNum("261571a1-4e16-11ed-7867-00505682218a"), defaultValue)
    var finUticajDrustvo = getItemAttributeValue(pObject, ArisData.ActiveFilter().UserDefinedAttributeTypeNum("34385ef1-4e16-11ed-7867-00505682218a"), defaultValue)

    //*************OBJ Mere**************************
    var objDefsMere = pObject.getConnectedObjs([Constants.OT_FUNC_INST], Constants.EDGES_INOUT)
    objDefsMere = ArisData.sort(objDefsMere, Constants.AT_NAME, Context.getSelectedLanguage())
    if (objDefsMere.length != 0) {
        for (var i = 0; i < objDefsMere.length; i++) {

            var meraGuid = objDefsMere[i].GUID()
            //var mereUpravljanjeRizicimaName = objDefsMere[i].Name(giLoc)           
            var mereUpravljanjeRizicimaName = getItemAttributeValue(objDefsMere[i], Constants.AT_DESC, defaultValue)
            var opisOcekivanogRez = getItemAttributeValue(objDefsMere[i], ArisData.ActiveFilter().UserDefinedAttributeTypeNum("74558751-4e17-11ed-7867-00505682218a"), defaultValue)
            var rokIspunjenaAktivnosti = getItemAttributeValue(objDefsMere[i], ArisData.ActiveFilter().UserDefinedAttributeTypeNum("88ef6c80-4e17-11ed-7867-00505682218a"), defaultValue)
            var odgovoran = getItemAttributeValue(objDefsMere[i], ArisData.ActiveFilter().UserDefinedAttributeTypeNum("9580bd51-4e17-11ed-7867-00505682218a"), defaultValue)
            var troskoviAktivnosti = getItemAttributeValue(objDefsMere[i], ArisData.ActiveFilter().UserDefinedAttributeTypeNum("a983b9b1-4e17-11ed-7867-00505682218a"), defaultValue)
            var opisRealizacijeMera = getItemAttributeValue(objDefsMere[i], ArisData.ActiveFilter().UserDefinedAttributeTypeNum("baa763e0-4e17-11ed-7867-00505682218a"), defaultValue)
            var statusRealizacijeMera = getItemAttributeValue(objDefsMere[i], ArisData.ActiveFilter().UserDefinedAttributeTypeNum("cb0fd461-4e17-11ed-7867-00505682218a"), defaultValue)


            printRow(rskObjectName, rskDescription, faktorRizikaPrilike, posledicaRizika, orgDeo, pozicijaVlasnik, imePrezimeVlasnik, revidiranjeIzvrsio, katBiznisProcesa, proces, grupaRizika, podgrupaRizika, pristupProceni, koristenaMetoda, verovatnocaMetoda1, verovatnocaMetoda2, verovatnocaMetoda3, uticajMetoda1, uticajMetoda2, uticajMetoda3, finUticaj,
                nivo, odgNaRizikPriliku, mereUpravljanjeRizicimaName, opisOcekivanogRez, rokIspunjenaAktivnosti, odgovoran, troskoviAktivnosti, opisRealizacijeMera, statusRealizacijeMera, rizikRealizacija, opisKalkulacijeOstvarenja, finUticajDrustvo, opisRealizacijeRizika, pokazateljRizika, uticajRealizacijeRizika, rskGuid, meraGuid, pExcelWorkBook, rowColorBr);

            //rizikRealizacija, opisRealizacijeRizika, pokazateljRizika, uticajRealizacijeRizika, mereUpravljanjeRizicimaName, opisOcekivanogRez, rokIspunjenaAktivnosti, odgovoran, troskoviAktivnosti, opisRealizacijeMera, statusRealizacijeMera,

            funcBr++;
            rowColorBr++
        }
        //**********************************************
    } else {
        printRow(rskObjectName, rskDescription, faktorRizikaPrilike, posledicaRizika, orgDeo, pozicijaVlasnik, imePrezimeVlasnik, revidiranjeIzvrsio, katBiznisProcesa, proces, grupaRizika, podgrupaRizika, pristupProceni, koristenaMetoda, verovatnocaMetoda1, verovatnocaMetoda2, verovatnocaMetoda3, uticajMetoda1, uticajMetoda2, uticajMetoda3, finUticaj,
            nivo, odgNaRizikPriliku, "", "", "", "", "", "", "", rizikRealizacija, opisKalkulacijeOstvarenja, finUticajDrustvo, opisRealizacijeRizika, pokazateljRizika, uticajRealizacijeRizika, rskGuid, "", pExcelWorkBook, rowColorBr);

        funcBr++;
        rowColorBr++
    }

    //merge cells after print
    mergeCells(funcBr, pExcelWorkBook)

}

function getItemAttributeValue(pArisItem, piAttributeTypeNum, sDefaultValue) {
    var sValue = "";
    try {
        var lAttObj = pArisItem.Attribute(piAttributeTypeNum, giLoc);
        if (lAttObj.IsMaintained() == true) {
            sValue = lAttObj.getValue();
        } else {
            sValue = sDefaultValue;
        }
    } catch (ex) {
        Dialogs.MsgBox("Error : " + ex);
    }
    return sValue;
}

function ObjsListNamesToString(paObjList) {
    var psResult = "";
    var liCounter = 0;
    var nrOfElements = paObjList.length;
    for each(loObj in paObjList) {
        if (liCounter + 1 == nrOfElements) {
            psResult += gsBulletSign + " " + loObj.Name(giLoc);
        } else {
            psResult += gsBulletSign + " " + loObj.Name(giLoc) + "\n";
        }
    }
    return psResult;
}

// remove line break from String parameter psText
function filterCarriageReturn(psText) {
    var lsResult = "";
    if (psText != null && psText.equalsIgnoreCase("") == false) {
        lsResult = psText.replaceAll("\\r|\\n", "");
    }
    return lsResult;
}

function mergeCells(funcBr, pExcelWorkBook) {
    if (funcBr > 1) {
        var regionSheet = pExcelWorkBook.getSheetAt(0);
        var fromRow = giRowCounter - funcBr

        regionSheet.addRegion(0, 0, fromRow, giRowCounter - 1)
        regionSheet.addRegion(1, 1, fromRow, giRowCounter - 1)
        regionSheet.addRegion(2, 2, fromRow, giRowCounter - 1)

        regionSheet.addRegion(3, 3, fromRow, giRowCounter - 1)
        regionSheet.addRegion(4, 4, fromRow, giRowCounter - 1)
        regionSheet.addRegion(5, 5, fromRow, giRowCounter - 1)
        regionSheet.addRegion(6, 6, fromRow, giRowCounter - 1)
        regionSheet.addRegion(7, 7, fromRow, giRowCounter - 1)

        regionSheet.addRegion(8, 8, fromRow, giRowCounter - 1)
        regionSheet.addRegion(9, 9, fromRow, giRowCounter - 1)
        regionSheet.addRegion(10, 10, fromRow, giRowCounter - 1)
        regionSheet.addRegion(11, 11, fromRow, giRowCounter - 1)
        regionSheet.addRegion(12, 12, fromRow, giRowCounter - 1)
        regionSheet.addRegion(13, 13, fromRow, giRowCounter - 1)
        regionSheet.addRegion(14, 14, fromRow, giRowCounter - 1)
        regionSheet.addRegion(15, 15, fromRow, giRowCounter - 1)
        regionSheet.addRegion(16, 16, fromRow, giRowCounter - 1)
        regionSheet.addRegion(17, 17, fromRow, giRowCounter - 1)
        regionSheet.addRegion(18, 18, fromRow, giRowCounter - 1)
        regionSheet.addRegion(19, 19, fromRow, giRowCounter - 1)
        regionSheet.addRegion(20, 20, fromRow, giRowCounter - 1)
        regionSheet.addRegion(21, 21, fromRow, giRowCounter - 1)
        //regionSheet.addRegion(22, 22, fromRow, giRowCounter - 1)

        regionSheet.addRegion(30, 30, fromRow, giRowCounter - 1)
        regionSheet.addRegion(31, 31, fromRow, giRowCounter - 1)
        regionSheet.addRegion(32, 32, fromRow, giRowCounter - 1)
        regionSheet.addRegion(33, 33, fromRow, giRowCounter - 1)
        regionSheet.addRegion(34, 34, fromRow, giRowCounter - 1)
        regionSheet.addRegion(35, 35, fromRow, giRowCounter - 1)
        regionSheet.addRegion(36, 36, fromRow, giRowCounter - 1)
    }
}

function printRow(rskObjectName, rskDescription, faktorRizikaPrilike, posledicaRizika, orgDeo, pozicijaVlasnik, imePrezimeVlasnik, revidiranjeIzvrsio, katBiznisProcesa, proces, grupaRizika, podgrupaRizika, pristupProceni, koristenaMetoda, verovatnocaMetoda1, verovatnocaMetoda2, verovatnocaMetoda3, uticajMetoda1, uticajMetoda2, uticajMetoda3, finUticaj,
    nivo, odgNaRizikPriliku, mereUpravljanjeRizicimaName, opisOcekivanogRez, rokIspunjenaAktivnosti, odgovoran, troskoviAktivnosti, opisRealizacijeMera, statusRealizacijeMera, rizikRealizacija, opisKalkulacijeOstvarenja, finUticajDrustvo, opisRealizacijeRizika, pokazateljRizika, uticajRealizacijeRizika, rskGuid, meraGuid, pExcelWorkBook, rowColorBr) {
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
    lCell.setCellValue(rskObjectName);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(1);
    lCell.setCellValue(rskDescription);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(2);
    lCell.setCellValue(faktorRizikaPrilike);
    lCell.setCellStyle(crDataRowStyle);

    // Input Name
    var lCell = lModelRow.createCell(3);
    lCell.setCellValue(posledicaRizika);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(4);
    lCell.setCellValue(orgDeo);
    lCell.setCellStyle(crDataRowStyle);

    //var lCell = lModelRow.createCell(5);
    //lCell.setCellValue(pozicijaVlasnik);
    //lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(6-1);
    lCell.setCellValue(imePrezimeVlasnik);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(7-1);
    lCell.setCellValue(revidiranjeIzvrsio);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(8-1);
    lCell.setCellValue(katBiznisProcesa);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(9-1);
    lCell.setCellValue(proces);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(10-1);
    lCell.setCellValue(grupaRizika);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(11-1);
    lCell.setCellValue(podgrupaRizika);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(12-1);
    lCell.setCellValue(pristupProceni);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(13-1);
    lCell.setCellValue(koristenaMetoda);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(14-1);
    lCell.setCellValue(verovatnocaMetoda1);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(15-1);
    lCell.setCellValue(verovatnocaMetoda2);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(16-1);
    lCell.setCellValue(verovatnocaMetoda3);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(17-1);
    lCell.setCellValue(uticajMetoda1);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(18-1);
    lCell.setCellValue(uticajMetoda2);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(19-1);
    lCell.setCellValue(uticajMetoda3);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(20-1);
    lCell.setCellValue(finUticaj.toString());
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(21-1);
    lCell.setCellValue(nivo);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(22-1);
    lCell.setCellValue(odgNaRizikPriliku);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(23-1);
    lCell.setCellValue(mereUpravljanjeRizicimaName);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(24-1);
    lCell.setCellValue(opisOcekivanogRez);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(25-1);
    lCell.setCellValue(rokIspunjenaAktivnosti);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(26-1);
    lCell.setCellValue(odgovoran);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(27-1);
    lCell.setCellValue(troskoviAktivnosti);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(28-1);
    lCell.setCellValue(opisRealizacijeMera);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(29-1);
    lCell.setCellValue(statusRealizacijeMera);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(30-1);
    lCell.setCellValue(rizikRealizacija);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(31-1);
    lCell.setCellValue(opisKalkulacijeOstvarenja);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(32-1);
    lCell.setCellValue(finUticajDrustvo);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(33-1);
    lCell.setCellValue(opisRealizacijeRizika);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(34-1);
    lCell.setCellValue(pokazateljRizika);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(35-1);
    lCell.setCellValue(uticajRealizacijeRizika);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(36-1);
    lCell.setCellValue(rskGuid);
    lCell.setCellStyle(iStyleGuid);

    var lCell = lModelRow.createCell(37-1);
    lCell.setCellValue(meraGuid);
    lCell.setCellStyle(iStyleGuid);


    // up counter
    giRowCounter++;
    //} catch (ex) {
    //   Dialogs.MsgBox("Error : " + ex);
    // }
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

main()