var epcStartModels = ArisData.getSelectedModels()
var startObjectsOcc = ArisData.getSelectedObjOccs()
var nLocale = Context.getSelectedLanguage()

var gs_xlsTemplate = "testCaseTemplate.xlsx"
var go_dataFromTemplate = Context.getFile(gs_xlsTemplate, Constants.LOCATION_SCRIPT) //spremi dokument u ovu varijablu
var go_excel = createNewExcelWorkBook() //kreiraj excell book 

var xlsSheet = go_excel.getSheetAt(0) // dohvati prvi sheet na kojem se nalazi template
go_xlsStyle = new setXlsStyles() //za font

var targetGroup = "" //varijabla za odabranu grupi
var cancelGroupDialog = false //varijabla za provjeru cancel buttona group dialog
var modelName = "" //varijabla za ime modela
var cancelSetModelNameDialog = false //varijabla za provjeru cancel buttona setModel name dialog

initializeGroupDialog() //prikazi group dialog

if (cancelGroupDialog != true) { //ako nije pritisnut cancel na group dialogu prikazi set model name dialog
	setModelNameDialog()
}

//ako nije pritisnuto cancel dugme pokreni main
if (cancelGroupDialog != true && cancelSetModelNameDialog != true) {
	main() //main function
}
else {
	//ne prikazuj outpufile, ne prikazuj success poruku
	Context.setProperty(Constants.PROPERTY_SHOW_OUTPUT_FILE, false)
	Context.setProperty(Constants.PROPERTY_SHOW_SUCCESS_MESSAGE, false)
}

//*************************************************************************************************************************************************************************
//                                                                               FUNKCIJE
//*************************************************************************************************************************************************************************

function getStartFunctions() {
	if (startObjectsOcc.length != 0) {
		var tempFunctionObjects = startObjectsOcc
	}
	else if (epcStartModels.length != 0) {
		var epcModel = epcStartModels[0]

		//dohvati sve funkcije sa modela 
		var tempFunctionObjects = epcModel.ObjOccListFilter(Constants.OT_FUNC) //dohvati sve occ objekte funkcija sa modela     
	}

	return tempFunctionObjects
}

function main() {
	var tempFunctionObjects = getStartFunctions() // dohvati funkcije na kojima se starta report

	//Izbaci Interface objekte iz niza funkcija i rezultat spremi u varijablu functionObjectsOcc
	var functionObjectsOcc = new Array()
	for (var j = 0; j < tempFunctionObjects.length; j++) {
		if (tempFunctionObjects[j].OrgSymbolNum() != Constants.ST_PRCS_IF) {
			functionObjectsOcc.push(tempFunctionObjects[j])
		}
	}

	//dohvati samo funkcije kojima je atribut Test case relevant postavljen u TRUE
	var testRelevantFunctions = new Array()
	for (var j = 0; j < functionObjectsOcc.length; j++) {
		var testCaseRelevant = functionObjectsOcc[j].ObjDef().Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum(AT_TEST_CASE_RELEVANT), nLocale).GetValue(false)
		if (testCaseRelevant == "True") {
			testRelevantFunctions.push(functionObjectsOcc[j])
		}
	}

	//Sortiraj funkcije prema atributu na funkciji "Function sequence number"
	testRelevantFunctions = sortBySequenceNumber(testRelevantFunctions)

	if (testRelevantFunctions.length > 0) { // ***samo ako postoje oznacene funkcije nastavi sa izvrsavanjem koda***

		//*******************************************************************************************************************
		//DOHVATI EXCEL TEMPLATE
		//Excel Template
		//prebaceno na pocetak skripte
		//*******************************************************************************************************************

		var rowNumber = 14 //prvi red u tablici za punjenje podataka 

		//idi kroz test case relevant funkcije i za svaku uzmi potrebne podatke, zatim taj podatak ubaci u tablicu
		for (var j = 0; j < testRelevantFunctions.length; j++) {
			//function name
			var functionName = testRelevantFunctions[j].ObjDef().Name(nLocale)

			//connected role
			var connectedRoles = getConnectedObjectsFromObjectOcc(testRelevantFunctions[j], Constants.OT_PERS_TYPE)
			var connectedRolesName = ""
			for (var k = 0; k < connectedRoles.length; k++) {
				if (k < connectedRoles.length - 1) {
					connectedRolesName = connectedRolesName + connectedRoles[k].ObjDef().Name(nLocale) + "\n"
				}
				else {
					connectedRolesName = connectedRolesName + connectedRoles[k].ObjDef().Name(nLocale)
				}
			}

			//connected screen
			var connectedScreens = getConnectedObjectsFromObjectOcc(testRelevantFunctions[j], Constants.OT_SCRN)

			//connected input data
			var inputObjects = testRelevantFunctions[j].getConnectedObjOccs(null, Constants.EDGES_IN) //dohvati sve objekte spojene ulaznom vezom na objektOcc funkcije 
			//idi kroz varijablu "inputObjects" i u varijablu "objects" spremi sve objekte tipa "ERM attribute"
			var inputErmAtt = new Array()
			for (var k = 0; k < inputObjects.length; k++) {
				if (inputObjects[k].ObjDef().TypeNum() == Constants.OT_ERM_ATTR) {
					inputErmAtt.push(inputObjects[k])
				}
			}

			//************************************************************************
			//erm atr value treba jos loviti **********!!!!!!!!!!!!!!!!!!!************
			//************************************************************************

			//ispisi screen za funkciju  nakon toga input objekte
			for (var k = 0; k < connectedScreens.length; k++) { //ovaj je za ispis redova za ERM att napraviti iznad isti samo za ispis rola

				xlsSheet.createRow(rowNumber).createCell(0).setCellValue(functionName) //naziv funkcije
				xlsSheet.createRow(rowNumber).createCell(0).setCellStyle(go_xlsStyle.FN_CELL_CUSTOM)

				xlsSheet.createRow(rowNumber).createCell(1).setCellValue(connectedScreens[k].ObjDef().Name(nLocale)) //naziv screena
				xlsSheet.createRow(rowNumber).createCell(1).setCellStyle(go_xlsStyle.FN_CELL_CUSTOM)

				xlsSheet.createRow(rowNumber).createCell(3).setCellValue(connectedRolesName)
				xlsSheet.createRow(rowNumber).createCell(3).setCellStyle(go_xlsStyle.FN_CELL_CUSTOM) //ispisi ih sve role jednu ispod druge, to je taj style
				rowNumber++ //povecaj brojac za novi red   
			}

			//kreiraj tablicu
			for (var k = 0; k < inputErmAtt.length; k++) { //ovaj je za ispis redova za ERM att napraviti iznad isti samo za ispis rola
				xlsSheet.createRow(rowNumber).createCell(0).setCellValue(functionName) //naziv funkcije
				xlsSheet.createRow(rowNumber).createCell(0).setCellStyle(go_xlsStyle.FN_CELL_CUSTOM)

				xlsSheet.createRow(rowNumber).createCell(2).setCellValue(inputErmAtt[k].ObjDef().Name(nLocale)) //input ERM att
				xlsSheet.createRow(rowNumber).createCell(2).setCellStyle(go_xlsStyle.FN_CELL_CUSTOM)

				rowNumber++ //povecaj brojac za novi red   
			}

		}

		//******************
		//CREATE NEW MODEL WITH FUNCTIONS OCC FROM THIS ONE. Izbornik di se odabire direktorij za spremanje model. Input box za upis imena modela ako ime vec postoji u grupi unesi drugo.
		//Kreiraj model u odabranoj grupi. Kreiraj occ kopije funkcija i spoji ih bez eventa. Takodjer na svakoj funkciji kreiraj i ostale potrebne podatke
		//dialozi pozvani prije main funkcije, zbog cancel buttona!!

		//kreiraj novi model u odabranoj grupi
		var newModel = targetGroup.CreateModel(Constants.MT_EEPC, modelName, nLocale) //kreiraj novi model u radnoj grupi

		//popuni kreirani model sa objektima funkcija
		var CX = 300 //get x coordinate
		var CY = 300 //get y coodridante  CY povecavaj za 300 tako da skace u novi red funkcija svaka

		var newOccurrence = new Array() //niz za spremanje novih kreiranih Occ kopija u novom modelu

		//kreiraj prvu funkciju u modelu
		newOccurrence.push(newModel.createObjOcc(Constants.ST_FUNC, testRelevantFunctions[0].ObjDef(), CX, CY, true, false)) //kreiraj prvi occ funkcije
		CY = CY + 300

		//idi kroz preostale funkcije (prva je vec kreirana)
		for (var j = 1; j < testRelevantFunctions.length; j++) {
			newOccurrence.push(newModel.createObjOcc(Constants.ST_FUNC, testRelevantFunctions[j].ObjDef(), CX, CY, true, false)) //kreiraj occ funkcije

			var sourceObject = newOccurrence[newOccurrence.length - 2] //novo kreirani predzadnji objekt
			var targetObject = newOccurrence[newOccurrence.length - 1] //novo kreirani zadnji objekt

			//ovo je dio koji definira točke na kojoj ce veze na objekt u modelu spojiti
			var sourcePoint       
			var targetPoint                
			var import_custom = JavaImporter(Packages.java.awt); // Java Package Import		        

			with(import_custom) {            
				sourcePoint = new Point(sourceObject.X() + sourceObject.Width() / 2, sourceObject.Y() + sourceObject.Height())
				targetPoint = new Point(targetObject.X() + targetObject.Width() / 2, targetObject.Y())        
			}                
			var points = new Array()        
			points[0] = sourcePoint        
			points[1] = targetPoint

			//stvori vezu izmedju trenutno kreiranog objekta i onog objekta koji se nalazi prije njega
			var newConnection = newModel.CreateCxnOcc(sourceObject, targetObject, Constants.CT_IS_PREDEC_OF_1, points, false, true).applyTemplate() //kreiraj vezu

			CY = CY + 300
		}
		//***************************************

		go_excel.write() //ispisi report output file

	}
	else
		Dialogs.MsgBox("There is no selected functions for test case!\n\nPlease first select functions then run report! ", Constants.MSGBOX_ICON_ERROR | Constants.MSGBOX_BTN_OK, "Error")
}

function createNewExcelWorkBook() {
	go_excel = null;
	try {
		var go_excel = Context.createExcelWorkbook(Context.getSelectedFile(), go_dataFromTemplate);
	}
	catch (e) {
		return gb_stop = true;
	}

	return go_excel;
}

function setModelNameDialog() {
	//dohvati ime za novi model. Ako ime postoji u odabranoj grupi ponovno unesi ime modela        
	var flag = false
	while (flag != true && cancelSetModelNameDialog == false) {
		modelName = Dialogs.InputBox("Enter model name", "Name input", "no name") //izbaci dialog
		if (modelName == "") { //ako pritisne na cancel zavrsi report
			//code here
			cancelSetModelNameDialog = true
		}
		else {
			var modelsInGroup = targetGroup.ModelList()
			var flag2 = false
			for (var j = 0; j < modelsInGroup.length; j++) {
				if (modelsInGroup[j].Name(nLocale) == modelName) {
					var flag2 = true
				}
			}
			if (flag2 == true) {
				Dialogs.MsgBox("Model with the same name already exist in selected group", Constants.MSGBOX_ICON_WARNING | Constants.MSGBOX_BTN_OK, "Warning")
			}
			else if (flag2 == false) {
				flag = true
			}
		}
	}
}

function initializeGroupDialog() {
	//dialog za odabir direktorija
	var serverName = ArisData.getActiveDatabase().ServerName()
	var dbName = ArisData.getActiveDatabase().Name(Context.getSelectedLanguage())
	var selection = Dialogs.BrowseArisItems("Select a group", "Groups", serverName, dbName, Constants.CID_GROUP, null, "") //OID grupe je spremljen tu
	if (selection == "") {
		cancelGroupDialog = true
	}
	else {
		//pomocu OID-a pronadji tocno o kojoj se grupi radi
		var database = ArisData.getActiveDatabase()
		targetGroup = database.FindOID(selection) //pronadji tocno koja je to grupa   
	}
}

function sortByPosition(oobjectocclist) {
	for (var i = 0; i < oobjectocclist.length - 1; i++)
		for (var j = i + 1; j < oobjectocclist.length; j++) {
			if (oobjectocclist[i].Y() > oobjectocclist[j].Y()) {
				var temp = oobjectocclist[i]
				oobjectocclist[i] = oobjectocclist[j]
				oobjectocclist[j] = temp
			}
			else if (oobjectocclist[i].Y() == oobjectocclist[j].Y() && oobjectocclist[i].X() > oobjectocclist[j].X()) {
				var temp = oobjectocclist[i]
				oobjectocclist[i] = oobjectocclist[j]
				oobjectocclist[j] = temp
			}
		}

	return oobjectocclist
}

function getConnectedObjectsFromObjectOcc(objectOcc, objectTypeNum) {
	//getConnectedObjectsFromObjectOcc(functionObjectOcc, Constants.OT_RULE)
	var connectedObjects = objectOcc.getConnectedObjOccs(null, Constants.EDGES_INOUT)
	var objects = new Array()

	for (var i = 0; i < connectedObjects.length; i++) {
		if (connectedObjects[i].ObjDef().TypeNum() == objectTypeNum) {
			objects.push(connectedObjects[i])
		}
	}

	return objects
}

function getObjectDefAtrValue(objectDef, atribut) {
	//getObjectDefAtrValue(objectDef[i], AT_IME_I_PREZIME_NA_POZICIJI)  
	if (objectDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum(atribut), nLocale).GetValue(false) != undefined) {
		var atrValue = objectDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum(atribut), nLocale).GetValue(false)
		return atrValue
	}
	else {
		var atrValue = ""
		return atrValue
	}
}

function sortBySequenceNumber(functionObjectsOcc) {
	//niz=[1.1.1, 5.1.2, 4.1.2, 2.1.3]  niz = sortVersions(niz)   
	functionObjectsOcc.sort(function(a, b) {
		var atribut_a = getObjectDefAtrValue(a.ObjDef(), AT_FUNCTION_SEQUENCE_NUMBER)
		var atribut_b = getObjectDefAtrValue(b.ObjDef(), AT_FUNCTION_SEQUENCE_NUMBER)

		var atribut_a = String(atribut_a)
		var atribut_b = String(atribut_b)

		var nums1 = atribut_a.split(".")
		var nums2 = atribut_b.split(".")

		for (var i = 0; i < nums1.length; i++) {
			if (nums2[i]) { // assuming 5..2 is invalid
				if (nums1[i] !== nums2[i]) {
					return nums1[i] - nums2[i]
				} // else continue
			}
			else {
				return 1 // no second number in b
			}
		}
		return -1 // was missing case b.len > a.len
	});

	return functionObjectsOcc
}

function RGB(r, g, b) {
	return (new java.awt.Color(r / 255.0, g / 255.0, b / 255.0, 1)).getRGB() & 0xFFFFFF
}

function setXlsStyles() {
	//colors
	var black = 0;
	var grey = 22;
	var white = 1;
	var red = 60; // Anubis 345052 - Dark red
	var green = 17;
	var orange = 52;
	var blue = 39;
	var yellow = 5;

	//fonts
	//font black normal
	this.FN = go_excel.createFont();
	this.FN.setFontName("Calibri");
	this.FN.setColor(black);
	//font black bold
	this.FNB = go_excel.createFont();
	this.FNB.setFontName("Arial");
	this.FNB.setColor(black);
	this.FNB.setBoldweight(0x2bc);
	//font grey normal
	this.FGREY = go_excel.createFont();
	this.FGREY.setFontName("Arial");
	this.FGREY.setColor(grey);
	//font grey bold
	this.FGREYB = go_excel.createFont();
	this.FGREYB.setFontName("Arial");
	this.FGREYB.setColor(grey);
	this.FGREYB.setBoldweight(0x2bc);
	//font red bold    
	this.FRB = go_excel.createFont();
	this.FRB.setFontName("Arial");
	this.FRB.setColor(red);
	this.FRB.setBoldweight(0x2bc);
	//font blue bold
	this.FBB = go_excel.createFont();
	this.FBB.setFontName("Arial");
	this.FBB.setColor(blue);
	this.FBB.setBoldweight(0x2bc);
	//font green bold
	this.FGB = go_excel.createFont();
	this.FGB.setFontName("Arial");
	this.FGB.setColor(green);
	this.FGB.setBoldweight(0x2bc);
	//font orange bold   
	this.FOB = go_excel.createFont();
	this.FOB.setFontName("Arial");
	this.FOB.setColor(orange);
	this.FOB.setBoldweight(0x2bc);
	//font white bold         
	this.FWB = go_excel.createFont();
	this.FWB.setFontName("Arial");
	this.FWB.setColor(white);
	this.FWB.setBoldweight(0x2bc);

	//cell complete styles
	this.basic = go_excel.createCellStyle(this.FN, 1, 1, 1, 1, Constants.C_GRAY, Constants.C_GRAY, Constants.C_GRAY, Constants.C_GRAY, Constants.ALIGN_LEFT, Constants.ALIGN_CENTER, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.SOLID_FOREGROUND);
	this.basic.setWrapText(true);

	//1-a  (font normal background normal)    
	this.FN_BGN = go_excel.createCellStyle(this.FN, 1, 1, 1, 1, Constants.C_GRAY, Constants.C_GRAY, Constants.C_GRAY, Constants.C_GRAY, Constants.ALIGN_LEFT, Constants.ALIGN_CENTER, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.SOLID_FOREGROUND);
	this.FN_BGN.setWrapText(true);
	//1-my  (font normal background normal no lines, centered)    
	this.FN1_BGN = go_excel.createCellStyle(this.FN, 1, 1, 1, 1, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_CENTER, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.SOLID_FOREGROUND);
	this.FN1_BGN.setWrapText(true);
	//1-my2  (font normal background normal no lines, centered)    
	this.FN2_BGN = go_excel.createCellStyle(this.FN, 1, 1, 1, 1, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_CENTER, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.SOLID_FOREGROUND);
	this.FN2_BGN.setWrapText(true);
	//1-b  (font normal background yellow)
	this.FN_BGY = go_excel.createCellStyle(this.FN, 1, 1, 1, 1, Constants.C_GRAY, Constants.C_GRAY, Constants.C_GRAY, Constants.C_GRAY, Constants.ALIGN_LEFT, Constants.ALIGN_CENTER, commonUtils.color.setColor(255, 255, 204), commonUtils.color.setColor(255, 255, 204), Constants.SOLID_FOREGROUND);
	this.FN_BGY.setWrapText(true);
	//1-c  (font normal background normal)    
	this.FNB_BGN = go_excel.createCellStyle(this.FNB, 1, 1, 1, 1, Constants.C_GRAY, Constants.C_GRAY, Constants.C_GRAY, Constants.C_GRAY, Constants.ALIGN_LEFT, Constants.ALIGN_CENTER, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.SOLID_FOREGROUND);
	this.FNB_BGN.setWrapText(true);
	//1-d  (font normal background yellow)
	this.FNB_BGY = go_excel.createCellStyle(this.FNB, 1, 1, 1, 1, Constants.C_GRAY, Constants.C_GRAY, Constants.C_GRAY, Constants.C_GRAY, Constants.ALIGN_LEFT, Constants.ALIGN_CENTER, commonUtils.color.setColor(255, 255, 204), commonUtils.color.setColor(255, 255, 204), Constants.SOLID_FOREGROUND);
	this.FNB_BGY.setWrapText(true);
	//2-a  (font red bold background normal)
	this.FRB_BGN = go_excel.createCellStyle(this.FRB, 1, 1, 1, 1, Constants.C_GREY_25_PERCENT, Constants.C_GREY_25_PERCENT, Constants.C_GREY_25_PERCENT, Constants.C_GREY_25_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_CENTER, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.SOLID_FOREGROUND);
	this.FRB_BGN.setWrapText(true);
	//2-b  (font red bold background yellow)
	this.FRB_BGY = go_excel.createCellStyle(this.FRB, 1, 1, 1, 1, Constants.C_GREY_25_PERCENT, Constants.C_GREY_25_PERCENT, Constants.C_GREY_25_PERCENT, Constants.C_GREY_25_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_CENTER, commonUtils.color.setColor(255, 255, 204), commonUtils.color.setColor(255, 255, 204), Constants.SOLID_FOREGROUND);
	this.FRB_BGY.setWrapText(true);
	//3-a  (font blue bold background normal)
	this.FBB_BGN = go_excel.createCellStyle(this.FBB, 1, 1, 1, 1, Constants.C_GREY_25_PERCENT, Constants.C_GREY_25_PERCENT, Constants.C_GREY_25_PERCENT, Constants.C_GREY_25_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_CENTER, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.SOLID_FOREGROUND);
	this.FBB_BGN.setWrapText(true);
	//3-b  (font blue bold background yellow)
	this.FBB_BGY = go_excel.createCellStyle(this.FBB, 1, 1, 1, 1, Constants.C_GREY_25_PERCENT, Constants.C_GREY_25_PERCENT, Constants.C_GREY_25_PERCENT, Constants.C_GREY_25_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_CENTER, commonUtils.color.setColor(255, 255, 204), commonUtils.color.setColor(255, 255, 204), Constants.SOLID_FOREGROUND);
	this.FBB_BGY.setWrapText(true);
	//4-a  (font green bold background normal)
	this.FGB_BGN = go_excel.createCellStyle(this.FGB, 1, 1, 1, 1, Constants.C_GREY_25_PERCENT, Constants.C_GREY_25_PERCENT, Constants.C_GREY_25_PERCENT, Constants.C_GREY_25_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_CENTER, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.SOLID_FOREGROUND);
	this.FGB_BGN.setWrapText(true);
	//4-b  (font green bold background yellow)
	this.FGB_BGY = go_excel.createCellStyle(this.FGB, 1, 1, 1, 1, Constants.C_GREY_25_PERCENT, Constants.C_GREY_25_PERCENT, Constants.C_GREY_25_PERCENT, Constants.C_GREY_25_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_CENTER, commonUtils.color.setColor(255, 255, 204), commonUtils.color.setColor(255, 255, 204), Constants.SOLID_FOREGROUND);
	this.FGB_BGY.setWrapText(true);
	//5-a  (font orange bold background normal)
	this.FOB_BGN = go_excel.createCellStyle(this.FOB, 1, 1, 1, 1, Constants.C_GREY_25_PERCENT, Constants.C_GREY_25_PERCENT, Constants.C_GREY_25_PERCENT, Constants.C_GREY_25_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_CENTER, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.SOLID_FOREGROUND);
	this.FOB_BGN.setWrapText(true);
	//5-b  (font orange bold background yellow)
	this.FOB_BGY = go_excel.createCellStyle(this.FOB, 1, 1, 1, 1, Constants.C_GREY_25_PERCENT, Constants.C_GREY_25_PERCENT, Constants.C_GREY_25_PERCENT, Constants.C_GREY_25_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_CENTER, commonUtils.color.setColor(255, 255, 204), commonUtils.color.setColor(255, 255, 204), Constants.SOLID_FOREGROUND);
	this.FOB_BGY.setWrapText(true);
	//6-a  (font white bold background gray)
	this.FWB_BGG = go_excel.createCellStyle(this.FWB, 1, 1, 1, 1, Constants.C_GREY_25_PERCENT, Constants.C_GREY_25_PERCENT, Constants.C_GREY_25_PERCENT, Constants.C_GREY_25_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_CENTER, Constants.C_GRAY, Constants.C_GRAY, Constants.SOLID_FOREGROUND);
	this.FWB_BGG.setWrapText(true);
	//7-a  (font orange bold background yellow)
	this.FGREYB_BGN = go_excel.createCellStyle(this.FGREYB, 1, 1, 1, 1, Constants.C_GREY_25_PERCENT, Constants.C_GREY_25_PERCENT, Constants.C_GREY_25_PERCENT, Constants.C_GREY_25_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_CENTER, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.SOLID_FOREGROUND);
	this.FGREYB_BGN.setWrapText(true);

	//******
	this.FN_CELL_CUSTOM = go_excel.createCellStyle(this.FN, 1, 1, 1, 1, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_CENTER, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.SOLID_FOREGROUND);
	this.FN_CELL_CUSTOM.setWrapText(true);
}