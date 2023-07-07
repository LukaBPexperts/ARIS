var oOutput = Context.createOutputObject()
var xlTableActive = false
var nLocale = Context.getSelectedLanguage()
var model = ArisData.getSelectedModels()
var connectedModel = getConnectedModel(model[0])

oOutput.DefineF("Calibri", "Calibri", 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_JUSTIFY, 0, 0, 0, 0, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_DEFAULT"), getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VTOP, 0, 0, 0, 0, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_HEADING_1"), getString("ID_DEFAULT_FONT"), 18, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_TOCENTRY0, 0, 0, 4, 4, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_HEADING_2"), getString("ID_DEFAULT_FONT"), 14, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_ITALIC | Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_TOCENTRY1, 0, 0, 2, 2, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_TABLE_CONTENT"), getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VTOP, 0, 0, 0, 0, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_HEADING_3"), getString("ID_DEFAULT_FONT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_ITALIC | Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_TOCENTRY2, 0, 0, 1, 1, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_HEADING_4"), getString("ID_DEFAULT_FONT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_ITALIC | Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_TOCENTRY3, 0, 0, 0, 0, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_INFO"), getString("ID_DEFAULT_FONT"), 14, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER | Constants.FMT_VTOP, 0, 0, 1.76, 8.82, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_TABLE_HEAD"), getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER | Constants.FMT_VTOP, 0, 0, 0, 0, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_TITLE"), getString("ID_DEFAULT_FONT"), 21, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER | Constants.FMT_VTOP, 0, 0, 1.76, 8.82, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_HEADER_FOOTER"), getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VTOP, 0, 0, 0, 0, 0, 1)

setupOutputObject(oOutput)

oOutput.SetTitle(Context.getScriptInfo(Constants.SCRIPT_NAME))

firstPage()

main()

oOutput.WriteReport()

function setupOutputObject(outputObj) {
	outputObj.SetAutoTOCNumbering(true)
	globalHeader(outputObj)

	globalFooter(outputObj)
}

function globalHeader(outputObj) {

	outputObj.BeginHeader()
	if (Context.getSelectedFormat() != Constants.OUTEXCEL)
		outputObj.BeginTable(100, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_NOBORDER, 0)
	outputObj.TableRow()
	outputObj.ResetFrameStyle()
	outputObj.SetFrameStyle(Constants.FRAME_TOP, 0, 0)
	outputObj.SetFrameStyle(Constants.FRAME_LEFT, 0, 0)
	outputObj.SetFrameStyle(Constants.FRAME_BOTTOM, 0, 0)
	outputObj.SetFrameStyle(Constants.FRAME_RIGHT, 0, 0)
	outputObj.TableCell("", 65, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
	outputObj.BeginParagraph(Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
	var image = odabirLogo(model[0])
    if(image != undefined)
	outputObj.OutGraphic(image, -1, 74, 30)
	outputObj.EndParagraph()
	outputObj.TableCell("", 35, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_RIGHT, 0)
	outputObj.BeginTable(100, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_NOBORDER, 0)
	outputObj.TableRow()
	outputObj.ResetFrameStyle()
	outputObj.TableCell("", 53.13, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_RIGHT, 0)
	outputObj.BeginParagraph(Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)
	outputObj.OutputLn("Naziv:", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT, 0.71)
	outputObj.EndParagraph()
	outputObj.TableCell("", 46.88, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_RIGHT, 0)
	outputObj.BeginParagraph(Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)
	outputObj.OutputLn(model[0].Name(nLocale), getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT, 0.71)
	outputObj.EndParagraph()
	outputObj.TableRow()
	outputObj.ResetFrameStyle()
	outputObj.TableCell("", 53.13, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_RIGHT, 0)
	outputObj.BeginParagraph(Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)
	outputObj.OutputLn("Šifra:", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT, 0.71)
	outputObj.EndParagraph()
	outputObj.TableCell("", 46.88, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_RIGHT, 0)
	outputObj.BeginParagraph(Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)
	outputObj.OutputLn(getAtrValueFromModel(model[0], AT_SIFRA_DOKUMENTA), getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT, 0.71)
	outputObj.EndParagraph()
	outputObj.TableRow()
	outputObj.ResetFrameStyle()
	outputObj.TableCell("", 53.13, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_RIGHT, 0)
	outputObj.BeginParagraph(Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)
	outputObj.OutputLn("Verzija:", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT, 0.71)
	outputObj.EndParagraph()
	outputObj.TableCell("", 46.88, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_RIGHT, 0)
	outputObj.BeginParagraph(Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)
	outputObj.OutputLn(getAtrValueFromModel(model[0], AT_VERZIJA), getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT, 0.71)
	outputObj.EndParagraph()
	outputObj.TableRow()
	outputObj.ResetFrameStyle()
	outputObj.TableCell("", 53.13, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_RIGHT, 0)
	outputObj.BeginParagraph(Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)
	outputObj.OutputLn("Datum primene:", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT, 0.71)
	outputObj.EndParagraph()
	outputObj.TableCell("", 46.88, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_RIGHT, 0)
	outputObj.BeginParagraph(Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)
	outputObj.OutputLn(getAtrValueFromModel(model[0], AT_DATUM_PRIMENE), getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT, 0.71)
	outputObj.EndParagraph()
	if (Context.getSelectedFormat() != Constants.OUTEXCEL)
		outputObj.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
	else
		outputObj.TableRow()
	if (Context.getSelectedFormat() != Constants.OUTEXCEL)
		outputObj.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
	else
		outputObj.TableRow()
	outputObj.EndHeader()
}

function globalFooter(outputObj) {
	outputObj.BeginFooter()
	if (Context.getSelectedFormat() != Constants.OUTEXCEL)
		outputObj.BeginTable(100, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_NOBORDER, 0)
	outputObj.TableRow()
	outputObj.ResetFrameStyle()
	outputObj.SetFrameStyle(Constants.FRAME_TOP, 0, 0)
	outputObj.SetFrameStyle(Constants.FRAME_LEFT, 0, 0)
	outputObj.SetFrameStyle(Constants.FRAME_BOTTOM, 0, 0)
	outputObj.SetFrameStyle(Constants.FRAME_RIGHT, 0, 0)
	outputObj.TableCell("", 70, getString("ID_DEFAULT_FONT"), 8, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
	outputObj.BeginParagraph(Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
	outputObj.OutputLn(getAtrValueFromModel(model[0], AT_FOOTER), getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0.71)
	outputObj.EndParagraph()
	outputObj.TableCell("", 30, getString("ID_DEFAULT_FONT"), 8, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_RIGHT, 0)
	outputObj.BeginParagraph(Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)
	outputObj.Output("Strana ", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT, 0.71)
	outputObj.OutputField(Constants.FIELD_PAGE, getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT)
	outputObj.Output("/", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT, 0.71)
	outputObj.OutputField(Constants.FIELD_NUMPAGES, getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT)
	outputObj.OutputLn("", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT, 0.71)
	outputObj.EndParagraph()
	if (Context.getSelectedFormat() != Constants.OUTEXCEL)
		outputObj.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
	else
		outputObj.TableRow()
	outputObj.EndFooter()
}

function globalHeaderFirstPage(outputObj) {

	outputObj.BeginHeader()
	if (Context.getSelectedFormat() != Constants.OUTEXCEL)
		outputObj.BeginTable(100, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_NOBORDER, 0)
	outputObj.TableRow()
	outputObj.ResetFrameStyle()
	outputObj.SetFrameStyle(Constants.FRAME_TOP, 0, 0)
	outputObj.SetFrameStyle(Constants.FRAME_LEFT, 0, 0)
	outputObj.SetFrameStyle(Constants.FRAME_BOTTOM, 0, 0)
	outputObj.SetFrameStyle(Constants.FRAME_RIGHT, 0, 0)
	outputObj.TableCell("", 65, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
	outputObj.BeginParagraph(Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
	var image = odabirLogo(model[0])
     if(image != undefined)
	outputObj.OutGraphic(image, -1, 74, 30)
	outputObj.EndParagraph()
	outputObj.TableCell("", 35, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_RIGHT, 0)
	outputObj.BeginTable(100, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_NOBORDER, 0)
	outputObj.TableRow()
	outputObj.ResetFrameStyle()
	outputObj.TableCell("", 53.13, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_RIGHT, 0)
	outputObj.BeginParagraph(Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)
	outputObj.OutputLn("Šifra:", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT, 0.71)
	outputObj.EndParagraph()
	outputObj.TableCell("", 46.88, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_RIGHT, 0)
	outputObj.BeginParagraph(Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)
	outputObj.OutputLn(getAtrValueFromModel(model[0], AT_SIFRA_DOKUMENTA), getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT, 0.71)
	outputObj.EndParagraph()
	outputObj.TableRow()
	outputObj.ResetFrameStyle()
	outputObj.TableCell("", 53.13, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_RIGHT, 0)
	outputObj.BeginParagraph(Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)
	outputObj.OutputLn("Verzija:", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT, 0.71)
	outputObj.EndParagraph()
	outputObj.TableCell("", 46.88, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_RIGHT, 0)
	outputObj.BeginParagraph(Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)
	outputObj.OutputLn(getAtrValueFromModel(model[0], AT_VERZIJA), getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT, 0.71)
	outputObj.EndParagraph()
	outputObj.TableRow()
	outputObj.ResetFrameStyle()
	outputObj.TableCell("", 53.13, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_RIGHT, 0)
	outputObj.BeginParagraph(Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)
	outputObj.OutputLn("Datum primene:", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT, 0.71)
	outputObj.EndParagraph()
	outputObj.TableCell("", 46.88, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_RIGHT, 0)
	outputObj.BeginParagraph(Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)
	outputObj.OutputLn(getAtrValueFromModel(model[0], AT_DATUM_PRIMENE), getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT, 0.71)
	outputObj.EndParagraph()
	if (Context.getSelectedFormat() != Constants.OUTEXCEL)
		outputObj.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
	else
		outputObj.TableRow()
	if (Context.getSelectedFormat() != Constants.OUTEXCEL)
		outputObj.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
	else
		outputObj.TableRow()
	outputObj.EndHeader()
}

//**********************************************************************************************************************
//                                                     FUNKCIJE
//**********************************************************************************************************************

function RGB(r, g, b) {
	return (new java.awt.Color(r / 255.0, g / 255.0, b / 255.0, 1)).getRGB() & 0xFFFFFF
}

function firstPage() {
	var vlasnikProcesaPositionName = getPositionNameFromAssignedModel(model, [Constants.CT_IS_OWN])
	var vlasnikProcesaAtrName = getPositionAtrNameFromAssignedModel(model, [Constants.CT_IS_OWN])
	var izradioAtrName = getPositionAtrNameFromAssignedModel(model, [Constants.CT_IS_ALLOWED])
	var izradioName = getPositionNameFromAssignedModel(model, [Constants.CT_IS_ALLOWED])
	var odobrioAtrName = getPositionAtrNameFromAssignedModel(model, [Constants.CT_IS_GRANT_BY])
	var odobrioName = getPositionNameFromAssignedModel(model, [Constants.CT_IS_GRANT_BY])
	var verifikovaoAtrName = getPositionAtrNameFromAssignedModel(model, [Constants.CT_DEF_STATE])
	var verifikovaoName = getPositionNameFromAssignedModel(model, [Constants.CT_DEF_STATE])

	oOutput.BeginSection(false, Constants.SECTION_DEFAULT)

	oOutput.SetAutoTOCNumbering(true)

	globalHeaderFirstPage(oOutput)
	customFooter(oOutput)

	emptyrow(9)

	oOutput.BeginParagraph(Constants.FMT_CENTER, 0, 0, 0, 2, 0)
	oOutput.OutputLn("Procedura", getString("ID_DEFAULT_FONT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0)
	oOutput.OutputLn(model[0].Name(nLocale).toUpperCase(), getString("ID_DEFAULT_FONT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0)
	oOutput.EndParagraph()

	emptyrow(2)

	oOutput.BeginParagraph(Constants.FMT_CENTER, 0, 0, 0, 2, 0)
	oOutput.OutputLn("Vlasnik procesa", getString("ID_DEFAULT_FONT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0)
	oOutput.OutputLn(vlasnikProcesaPositionName, getString("ID_DEFAULT_FONT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0)
	oOutput.EndParagraph()

	emptyrow(1)

	oOutput.BeginParagraph(Constants.FMT_CENTER, 0, 0, 0, 0, 0)
	oOutput.OutputLn("Sektor/Služba", getString("ID_DEFAULT_FONT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0)
	oOutput.OutputLn(getOrganizationNameFromAssignedModel(model), getString("ID_DEFAULT_FONT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0)
	oOutput.EndParagraph()

	emptyrow(1)

	oOutput.EndSection()
	oOutput.BeginSection(false, Constants.SECTION_DEFAULT)
	globalHeader(oOutput)
	globalFooter(oOutput)

	oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
	oOutput.TableRow()
	oOutput.ResetFrameStyle()
	oOutput.TableCell("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255, 255, 204), 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn("ISTORIJA REVIZIJE", getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()

	oOutput.TableRow()
	oOutput.ResetFrameStyle()
	oOutput.TableCell("", (100 / 6) - 5, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255, 255, 204), 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn("Verzija", getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
	oOutput.TableCell("", (100 / 6) + 5, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255, 255, 204), 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn("Opis promene", getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
	oOutput.TableCell("", 100 / 6, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255, 255, 204), 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn("Autor", getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
	oOutput.TableCell("", 100 / 6, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255, 255, 204), 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn("Odobrio", getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
	oOutput.TableCell("", 100 / 6, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255, 255, 204), 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn("Verifikovao", getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
	oOutput.TableCell("", 100 / 6, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255, 255, 204), 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn("Datum puštanja u primenu", getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()

	var objectOccList = getObjectsOccList(connectedModel, Constants.OT_PACK)
   
    objectOccList = ArisData.sort( objectOccList, Constants.SORT_OBJDEF_NAME, Context.getSelectedLanguage() )

	for (var z = 0; z < objectOccList.length; z++) {

		oOutput.TableRow()
		oOutput.ResetFrameStyle()
		oOutput.TableCell("", (100 / 6) - 5, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_CENTER, 0)
		oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
		oOutput.OutputLn(objectOccList[z].ObjDef().Name(nLocale), getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
		oOutput.EndParagraph()
		oOutput.TableCell("", (100 / 6) + 5, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_CENTER, 0)
		oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
		oOutput.OutputLn(getObjectAtrValue(objectOccList[z], AT_OPIS_PROMENE_VERZIJE), getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
		oOutput.EndParagraph()
		oOutput.TableCell("", 100 / 6, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_CENTER, 0)
		oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
		oOutput.OutputLn(getObjectAtrValue(objectOccList[z], AT_AUTOR_VERZIJE), getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
		oOutput.EndParagraph()
		oOutput.TableCell("", 100 / 6, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_CENTER, 0)
		oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
		oOutput.OutputLn(getObjectAtrValue(objectOccList[z], AT_ODOBRIO_VERZIJU), getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
		oOutput.EndParagraph()
		oOutput.TableCell("", 100 / 6, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_CENTER, 0)
		oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
		oOutput.OutputLn(getObjectAtrValue(objectOccList[z], AT_VERIFIKOVAO_VERZIJU), getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
		oOutput.EndParagraph()
		oOutput.TableCell("", 100 / 6, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_CENTER, 0)
		oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
		oOutput.OutputLn(getObjectAtrValue(objectOccList[z], AT_DATUM_PUSTANJA_VERZIJE_U_PRIMENU), getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
		oOutput.EndParagraph()
	}
	oOutput.EndTable("", 100, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_VTOP, 0)

	ArisData.getActiveDatabase().clearCaches()

	emptyrow(3)

	oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
	oOutput.TableRow()
	oOutput.ResetFrameStyle()
	oOutput.TableCell("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255, 255, 204), 0, Constants.FMT_VTOP | Constants.FMT_CENTER, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn("VEZE SA DRUGIM DOKUMENTIMA", getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()

	oOutput.TableRow()
	oOutput.ResetFrameStyle()
	oOutput.TableCell("", 50, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255, 255, 204), 0, Constants.FMT_VTOP | Constants.FMT_CENTER, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn("Šifra dokumenta", getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
	oOutput.TableCell("", 50, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255, 255, 204), 0, Constants.FMT_VTOP | Constants.FMT_CENTER, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn("Naziv dokumenta", getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()

	var interfacesOccList = model[0].ObjOccListBySymbol([Constants.ST_PRCS_IF])

	sortByPosition(interfacesOccList)

	var nizmodela = new Array();

	for (var v = 0; v < interfacesOccList.length; v++) {
		var assignedModelToInterfaceList = interfacesOccList[v].ObjDef().AssignedModels(null)

		if (assignedModelToInterfaceList[0]) {
            nizmodela.push(assignedModelToInterfaceList[0])
		}
	}

	nizmodela = ArisData.Unique(nizmodela)

	for (var xyz = 0; xyz < nizmodela.length; xyz++) {
		oOutput.TableRow()
		oOutput.ResetFrameStyle()
		oOutput.TableCell("", 50, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_CENTER, 0)
		oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
		oOutput.OutputLn(getAtrValueFromModel(nizmodela[xyz], AT_SIFRA_DOKUMENTA), getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
		oOutput.EndParagraph()
		oOutput.TableCell("", 50, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
		oOutput.BeginParagraph(Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
		oOutput.OutputLn(nizmodela[xyz].Name(nLocale), getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0.71)
		oOutput.EndParagraph()
	}

	var documentsTehnicalTerm = model[0].ObjOccListFilter(Constants.OT_TECH_TRM)

	for (var dk = 0; dk < documentsTehnicalTerm.length; dk++) {
		oOutput.TableRow()
		oOutput.ResetFrameStyle()
		oOutput.TableCell("", 50, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_CENTER, 0)
		oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
		oOutput.OutputLn(documentsTehnicalTerm[dk].ObjDef().Attribute(Constants.AT_ID, nLocale).GetValue(false), getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
		oOutput.EndParagraph()
		oOutput.TableCell("", 50, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
		oOutput.BeginParagraph(Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
		oOutput.OutputLn(documentsTehnicalTerm[dk].ObjDef().Name(nLocale), getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0.71)
		oOutput.EndParagraph()
	}
	oOutput.EndTable("", 100, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0)

	oOutput.EndSection()

	ArisData.getActiveDatabase().clearCaches()
}

function emptyrow(number) {

	for (var i = 0; i < number; i++)
		oOutput.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
}

function getConnectedModel(model) {

	var objectOccProcess = model.ObjOccListFilter(Constants.OT_PRCS)

	var assignedModels = objectOccProcess[0].ObjDef().AssignedModels(null)

	return assignedModels[0]
}

function getPositionAtrNameFromAssignedModel(mainModel, string) {

	var connectedModel = getConnectedModel(mainModel[0])
	var objectOccFunction = connectedModel.ObjOccListFilter(Constants.OT_FUNC)
	var connectedDefPosition = objectOccFunction[0].ObjDef().getConnectedObjs([Constants.OT_POS], Constants.EDGES_INOUTASSIGN, string)
	var connectedDefPositionAtrName = connectedDefPosition[0].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum(AT_IME_I_PREZIME_NA_POZICIJI), nLocale).GetValue(false)

	return connectedDefPositionAtrName
}

function getPositionNameFromAssignedModel(mainModel, string) {

	var connectedModel = getConnectedModel(mainModel[0])
	var objectOccFunction = connectedModel.ObjOccListFilter(Constants.OT_FUNC)
	var connectedDefPosition = objectOccFunction[0].ObjDef().getConnectedObjs([Constants.OT_POS], Constants.EDGES_INOUTASSIGN, string)[0].Name(nLocale)

	return connectedDefPosition
}

function getOrganizationNameFromAssignedModel(mainModel) {

	var connectedModel = getConnectedModel(mainModel[0])
	var objectOccFunction = connectedModel.ObjOccListFilter(Constants.OT_ORG_UNIT)
	var connectedDefOrganizationName = objectOccFunction[0].ObjDef().Name(nLocale)
	return connectedDefOrganizationName
}

function getObjectAtrValue(objectOccList, atribut) {

	if (objectOccList.ObjDef().Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum(atribut), nLocale).GetValue(false) != undefined) {
		var atrValue = objectOccList.ObjDef().Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum(atribut), nLocale).GetValue(false)
		return atrValue
	}
	else {
		var atrValue = ""
		return atrValue
	}
}

function getObjectsOccList(model, objectType) {

	var objectOccList = model.ObjOccListFilter(objectType)

	return objectOccList
}

function getAtrValueFromModel(model, atribut) {

	var atribut_lista = model.AttrList(nLocale, [ArisData.ActiveFilter().UserDefinedAttributeTypeNum(atribut)])
	if (atribut_lista[0] != undefined) {
		var atribut_lista_value = atribut_lista[0].GetValue(false)
		return atribut_lista_value
	}
	else {
		var atribut_lista_value = ""
		return atribut_lista_value
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

function main() {

	oOutput.BeginParagraphF(getString("ID_STYLE_RD_HEADING_1"))
	oOutput.Output("Cilj i primena", getString("ID_STYLE_RD_DEFAULT"), 14, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
	oOutput.EndParagraph()
	oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
	oOutput.Output(getAtrValueFromModel(model[0], AT_CILJ_PRIMENE), getString("ID_STYLE_RD_DEFAULT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)
	oOutput.EndParagraph()

	oOutput.BeginParagraphF(getString("ID_STYLE_RD_HEADING_1"))
	oOutput.Output("Definicije i skraćenice", getString("ID_STYLE_RD_DEFAULT"), 14, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
	oOutput.EndParagraph()
	oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
	oOutput.Output(getAtrValueFromModel(model[0], AT_DEFINICIJE_I_SKRACENICE), getString("ID_STYLE_RD_DEFAULT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)
	oOutput.EndParagraph()

	oOutput.BeginParagraphF(getString("ID_STYLE_RD_HEADING_1"))
	oOutput.Output("Dijagram toka", getString("ID_STYLE_RD_DEFAULT"), 14, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
	oOutput.EndParagraph()
	oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
	oOutput.OutGraphic(model[0].Graphic(false, false, Context.getSelectedLanguage()), -1, 220, 220)
	oOutput.EndParagraph()

	ArisData.getActiveDatabase().clearCaches()

	oOutput.BeginParagraphF(getString("ID_STYLE_RD_HEADING_1"))
	oOutput.Output("Opis", getString("ID_STYLE_RD_DEFAULT"), 14, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
	oOutput.EndParagraph()

	var functionsOccList = model[0].ObjOccListBySymbol([Constants.ST_FUNC])

	sortByPosition(functionsOccList)

	for (var u = 0; u < functionsOccList.length; u++) {

		oOutput.BeginParagraphF(getString("ID_STYLE_RD_HEADING_2"))
		oOutput.Output(functionsOccList[u].ObjDef().Name(nLocale), getString("ID_STYLE_RD_DEFAULT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
		oOutput.EndParagraph()

		emptyrow(1)

		var connectedRoles = functionsOccList[u].ObjDef().getConnectedObjs([Constants.OT_PERS_TYPE], Constants.EDGES_TYP, new Array(Constants.CT_EXEC_1, Constants.CT_EXEC_2))

		var connectedPositionsListName = functionsOccList[u].ObjDef().getConnectedObjs(new Array(Constants.OT_POS, Constants.OT_PERS_TYPE), Constants.EDGES_TYP, new Array(Constants.CT_EXEC_1, Constants.CT_EXEC_2))

		oOutput.BeginTable(93, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT, 0)
		oOutput.TableRow()
		oOutput.ResetFrameStyle()
		oOutput.SetFrameStyle(Constants.FRAME_TOP, 0, 0)
		oOutput.SetFrameStyle(Constants.FRAME_LEFT, 0, 0)
		oOutput.SetFrameStyle(Constants.FRAME_BOTTOM, 0, 0)
		oOutput.SetFrameStyle(Constants.FRAME_RIGHT, 0, 0)
		oOutput.TableCell("", 17, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255, 255, 204), 0, Constants.FMT_VTOP | Constants.FMT_RIGHT, 0)
		oOutput.BeginParagraph(Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)
		oOutput.OutputLn("Odgovorna funkcija:", getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT, 0.71)
		oOutput.EndParagraph()
		oOutput.ResetFrameStyle()
		oOutput.SetFrameStyle(Constants.FRAME_TOP, 0, 0)
		oOutput.SetFrameStyle(Constants.FRAME_LEFT, 0, 0)
		oOutput.SetFrameStyle(Constants.FRAME_LEFT, 0, 0)
		oOutput.SetFrameStyle(Constants.FRAME_BOTTOM, 5, Constants.BRDR_DOT)
		oOutput.SetFrameStyle(Constants.FRAME_RIGHT, 0, 0)
		oOutput.TableCell("", 76, getString("ID_DEFAULT_FONT"), 9, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
		oOutput.BeginParagraph(Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
		for (var loop = 0; loop < connectedPositionsListName.length; loop++)
			oOutput.OutputLn(connectedPositionsListName[loop].Name(nLocale) + " ", getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0.71)
			//oOutput.OutputLn(getObjectDefListNamesString(connectedPositionsListName), getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0.71)
		oOutput.EndParagraph()
		oOutput.TableRow()
		oOutput.ResetFrameStyle()
		oOutput.SetFrameStyle(Constants.FRAME_TOP, 0, 0)
		oOutput.SetFrameStyle(Constants.FRAME_LEFT, 0, 0)
		oOutput.SetFrameStyle(Constants.FRAME_BOTTOM, 0, 0)
		oOutput.SetFrameStyle(Constants.FRAME_RIGHT, 0, 0)
		oOutput.TableCell("", 17, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255, 255, 204), 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
		oOutput.BeginParagraph(Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)
		oOutput.OutputLn("Rok:", getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0.71)
		oOutput.EndParagraph()
		oOutput.ResetFrameStyle()
		oOutput.SetFrameStyle(Constants.FRAME_TOP, 0, 0)
		oOutput.SetFrameStyle(Constants.FRAME_LEFT, 0, 0)
		oOutput.SetFrameStyle(Constants.FRAME_LEFT, 0, 0)
		oOutput.SetFrameStyle(Constants.FRAME_BOTTOM, 5, Constants.BRDR_DOT)
		oOutput.SetFrameStyle(Constants.FRAME_RIGHT, 0, 0)
		oOutput.TableCell("", 76, getString("ID_DEFAULT_FONT"), 9, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
		oOutput.BeginParagraph(Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
		oOutput.OutputLn(getObjectAtrValue(functionsOccList[u], AT_ROK_ZA_IZVRSENJE_AKTIVNOSTI), getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0.71)
		oOutput.EndParagraph()

		var connectedIncomingObjects = functionsOccList[u].ObjDef().getConnectedObjs(new Array(Constants.OT_INFO_CARR, Constants.OT_TECH_TRM), Constants.EDGES_IN, null)

		oOutput.TableRow()
		oOutput.ResetFrameStyle()
		oOutput.SetFrameStyle(Constants.FRAME_TOP, 0, 0)
		oOutput.SetFrameStyle(Constants.FRAME_LEFT, 0, 0)
		oOutput.SetFrameStyle(Constants.FRAME_BOTTOM, 0, 0)
		oOutput.SetFrameStyle(Constants.FRAME_RIGHT, 0, 0)
		oOutput.TableCell("", 17, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255, 255, 204), 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
		oOutput.BeginParagraph(Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)
		oOutput.OutputLn("Ulaz:", getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0.71)
		oOutput.EndParagraph()
		oOutput.ResetFrameStyle()
		oOutput.SetFrameStyle(Constants.FRAME_TOP, 0, 0)
		oOutput.SetFrameStyle(Constants.FRAME_LEFT, 0, 0)
		oOutput.SetFrameStyle(Constants.FRAME_LEFT, 0, 0)
		oOutput.SetFrameStyle(Constants.FRAME_BOTTOM, 5, Constants.BRDR_DOT)
		oOutput.SetFrameStyle(Constants.FRAME_RIGHT, 0, 0)
		oOutput.TableCell("", 76, getString("ID_DEFAULT_FONT"), 9, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
		oOutput.BeginParagraph(Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
		oOutput.OutputLn(getObjectDefListNamesString(connectedIncomingObjects), getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0.71)
		oOutput.EndParagraph()

		var connectedOutcomingObjects = functionsOccList[u].ObjDef().getConnectedObjs(new Array(Constants.OT_INFO_CARR, Constants.OT_ORG_UNIT, Constants.OT_ORG_UNIT_TYPE, Constants.OT_PERS, Constants.OT_POS, Constants.OT_PERS_TYPE, Constants.OT_TECH_TRM), Constants.EDGES_OUT, null)

		oOutput.TableRow()
		oOutput.ResetFrameStyle()
		oOutput.SetFrameStyle(Constants.FRAME_TOP, 0, 0)
		oOutput.SetFrameStyle(Constants.FRAME_LEFT, 0, 0)
		oOutput.SetFrameStyle(Constants.FRAME_BOTTOM, 0, 0)
		oOutput.SetFrameStyle(Constants.FRAME_RIGHT, 0, 0)
		oOutput.TableCell("", 17, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255, 255, 204), 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
		oOutput.BeginParagraph(Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)
		oOutput.OutputLn("Izlaz:", getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0.71)
		oOutput.EndParagraph()
		oOutput.ResetFrameStyle()
		oOutput.SetFrameStyle(Constants.FRAME_TOP, 0, 0)
		oOutput.SetFrameStyle(Constants.FRAME_LEFT, 0, 0)
		oOutput.SetFrameStyle(Constants.FRAME_LEFT, 0, 0)
		oOutput.SetFrameStyle(Constants.FRAME_BOTTOM, 5, Constants.BRDR_DOT)
		oOutput.SetFrameStyle(Constants.FRAME_RIGHT, 0, 0)
		oOutput.TableCell("", 76, getString("ID_DEFAULT_FONT"), 9, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
		oOutput.BeginParagraph(Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
		oOutput.OutputLn(getObjectDefListNamesString(connectedOutcomingObjects), getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0.71)
		oOutput.EndParagraph()

		var connectedAplicationObjects = functionsOccList[u].ObjDef().getConnectedObjs(new Array(Constants.OT_APPL_SYS, Constants.OT_APPL_SYS_TYPE), Constants.EDGES_INOUTASSIGN, null)

		oOutput.TableRow()
		oOutput.ResetFrameStyle()
		oOutput.SetFrameStyle(Constants.FRAME_TOP, 0, 0)
		oOutput.SetFrameStyle(Constants.FRAME_LEFT, 0, 0)
		oOutput.SetFrameStyle(Constants.FRAME_BOTTOM, 0, 0)
		oOutput.SetFrameStyle(Constants.FRAME_RIGHT, 0, 0)
		oOutput.TableCell("", 17, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255, 255, 204), 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
		oOutput.BeginParagraph(Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)
		oOutput.OutputLn("Aplikacije:", getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0.71)
		oOutput.EndParagraph()
		oOutput.ResetFrameStyle()
		oOutput.SetFrameStyle(Constants.FRAME_TOP, 0, 0)
		oOutput.SetFrameStyle(Constants.FRAME_LEFT, 0, 0)
		oOutput.SetFrameStyle(Constants.FRAME_LEFT, 0, 0)
		oOutput.SetFrameStyle(Constants.FRAME_BOTTOM, 5, Constants.BRDR_DOT)
		oOutput.SetFrameStyle(Constants.FRAME_RIGHT, 0, 0)
		oOutput.TableCell("", 76, getString("ID_DEFAULT_FONT"), 9, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
		oOutput.BeginParagraph(Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
		oOutput.OutputLn(getObjectDefListNamesString(connectedAplicationObjects), getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0.71)
		oOutput.EndParagraph()

		oOutput.EndTable("", 100, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0)

		emptyrow(1)

		oOutput.BeginParagraphF("Calibri")
		oOutput.Output(functionsOccList[u].ObjDef().Attribute(Constants.AT_DESC, nLocale).GetValue(false), "Calibri", 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_JUSTIFY, 0)
		oOutput.EndParagraph()

		emptyrow(1)
		getBlobDocumentFromObject(functionsOccList[u])

		emptyrow(2)
	}

	ArisData.getActiveDatabase().clearCaches()

	oOutput.BeginParagraphF(getString("ID_STYLE_RD_HEADING_1"))
	oOutput.Output("Merenja", getString("ID_STYLE_RD_DEFAULT"), 14, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
	oOutput.EndParagraph()

	var kpiOccList = getObjectsOccList(connectedModel, Constants.OT_KPI)

	kpiOccList.sort()

	for (var t = 0; t < kpiOccList.length; t++) {
		oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
		oOutput.OutputLn("      •  " + getObjectAtrValue(kpiOccList[t], AT_OPIS_KPI), getString("ID_STYLE_RD_DEFAULT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)
		oOutput.EndParagraph()
	}

	emptyrow(2)

	oOutput.BeginParagraphF(getString("ID_STYLE_RD_HEADING_1"))
	oOutput.Output("Zapisi", getString("ID_STYLE_RD_DEFAULT"), 14, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
	oOutput.EndParagraph()

	var documentsOccList = model[0].ObjOccListBySymbol(new Array(Constants.ST_DOC, Constants.ST_INFO_CARR_EDOC))

	documentsOccList = ArisData.sort(documentsOccList, Constants.SORT_OBJDEF_NAME, nLocale)
	var uniqueDocumentsOcc = Array()
	var objectF = ""
	var object_cntF = 0
	for (var a = 0; a < documentsOccList.length; a++) {
		var l_nameF = documentsOccList[a].ObjDef().Name(nLocale)
		if (l_nameF != objectF) {
			objectF = l_nameF;
			uniqueDocumentsOcc[object_cntF] = documentsOccList[a];
			object_cntF++;
		}
	}

	sortByPosition(uniqueDocumentsOcc)

	oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)
	oOutput.TableRow()
	oOutput.ResetFrameStyle()
	oOutput.SetFrameStyle(Constants.FRAME_BOTTOM, 0, 0)
	oOutput.TableCell("", 14.47 - 6, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255, 255, 204), 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
	oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
	oOutput.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
	oOutput.EndParagraph()
	oOutput.TableCell("", 14.26 + 6, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255, 255, 204), 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
	oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
	oOutput.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
	oOutput.EndParagraph()
	oOutput.ResetFrameStyle()
	oOutput.TableCell("", 71.3, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255, 255, 204), 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
	oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
	oOutput.OutputLn("Čuvanje", getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
	oOutput.TableRow()
	oOutput.ResetFrameStyle()
	oOutput.SetFrameStyle(Constants.FRAME_TOP, 0, 0)
	oOutput.TableCell("", 14.47 - 6, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255, 255, 204), 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
	oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
	oOutput.OutputLn("R.Br", getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
	oOutput.TableCell("", 14.26 + 6, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255, 255, 204), 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
	oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
	oOutput.OutputLn("Naziv Zapisa", getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
	oOutput.TableCell("", 14.26, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255, 255, 204), 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
	oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
	oOutput.OutputLn("Šifra formulara", getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
	oOutput.ResetFrameStyle()
	oOutput.SetFrameStyle(Constants.FRAME_TOP, 15, Constants.BRDR_NORMAL)
	oOutput.TableCell("", 13.84, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255, 255, 204), 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
	oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
	oOutput.OutputLn("Mesto", getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
	oOutput.TableCell("", 14.68, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255, 255, 204), 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
	oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
	oOutput.OutputLn("Način", getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
	oOutput.TableCell("", 14.26, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255, 255, 204), 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
	oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
	oOutput.OutputLn("Rok", getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
	oOutput.TableCell("", 14.26, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255, 255, 204), 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
	oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
	oOutput.OutputLn("Odgovorna funkcija za čuvanje dokumenata", getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()

	for (var s = 0; s < uniqueDocumentsOcc.length; s++) {
		oOutput.TableRow()
		oOutput.ResetFrameStyle()
		oOutput.TableCell("", 14.47 - 6, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
		oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
		oOutput.OutputLn(s + 1, getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)

		oOutput.EndParagraph()
		oOutput.TableCell("", 14.26 + 6, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
		oOutput.BeginParagraph(Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
		oOutput.OutputLn(uniqueDocumentsOcc[s].ObjDef().Name(nLocale), getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
		oOutput.EndParagraph()
		oOutput.TableCell("", 14.26, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
		oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
		oOutput.OutputLn(getObjectAtrValue(uniqueDocumentsOcc[s], AT_SIFRA_FORMULARA), getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0.71)
		oOutput.EndParagraph()
		oOutput.TableCell("", 13.84, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
		oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
		oOutput.OutputLn(getObjectAtrValue(uniqueDocumentsOcc[s], AT_MESTO_ARHIVIRANJA_ZAPISA), getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0.71)
		oOutput.EndParagraph()
		oOutput.TableCell("", 14.68, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
		oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
		oOutput.OutputLn(getObjectAtrValue(uniqueDocumentsOcc[s], AT_NACIN_ARHIVIRANJA_ZAPISA), getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0.71)
		oOutput.EndParagraph()
		oOutput.TableCell("", 14.26, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
		oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
		oOutput.OutputLn(getObjectAtrValue(uniqueDocumentsOcc[s], AT_PERIOD_ARHIVIRANJA_ZAPISA), getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0.71)
		oOutput.EndParagraph()
		oOutput.TableCell("", 14.26, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
		oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
		oOutput.OutputLn(getObjectAtrValue(uniqueDocumentsOcc[s], AT_ODGOVORNA_FUNKCIJA_ZA_ARHIVIRANJE_ZAPISA), getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0.71)
		oOutput.EndParagraph()
	}
	oOutput.EndTable("", 100, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0)

	ArisData.getActiveDatabase().clearCaches()
}

function customFooter(outputObj) {

	outputObj.BeginFooter()

	oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER | Constants.FMT_REPEAT_HEADER, 0)
	oOutput.TableRow()
	oOutput.ResetFrameStyle()
	oOutput.TableCell("", 33, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_CENTER, 0)

	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn(getPositionAtrNameFromAssignedModel(model, [Constants.CT_IS_ALLOWED]), getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.OutputLn(getPositionNameFromAssignedModel(model, [Constants.CT_IS_ALLOWED]), getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
	oOutput.TableCell("", 33, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn(getPositionAtrNameFromAssignedModel(model, [Constants.CT_IS_GRANT_BY]), getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.OutputLn(getPositionNameFromAssignedModel(model, [Constants.CT_IS_GRANT_BY]), getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
	oOutput.TableCell("", 33, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_CENTER, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn(getPositionAtrNameFromAssignedModel(model, [Constants.CT_DEF_STATE]), getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.OutputLn(getPositionNameFromAssignedModel(model, [Constants.CT_DEF_STATE]), getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()

	oOutput.TableRow()
	oOutput.ResetFrameStyle()
	oOutput.TableCell("", 33, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255, 255, 204), 0, Constants.FMT_VTOP | Constants.FMT_CENTER, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn("Izradio", getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
	oOutput.TableCell("", 33, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255, 255, 204), 0, Constants.FMT_VTOP | Constants.FMT_CENTER, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn("Odobrio", getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
	oOutput.TableCell("", 33, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255, 255, 204), 0, Constants.FMT_VTOP | Constants.FMT_CENTER, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn("Verifikovao", getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
	oOutput.EndTable("", 100, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_VTOP, 0)
	outputObj.EndFooter()
}

function getBlobDocumentFromObject(object) {
	var atrDocument = object.ObjDef().AttrList(nLocale, [Constants.AT_DOCUMENTATION])

	if (atrDocument[0] != undefined) {
		var dokument = atrDocument[0].getExternalDocument().getDocument()
		try {
			var docObj = Context.getOutputObject("tmp.doc", dokument)
			oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			oOutput.InsertDocument(docObj, false, false)
			oOutput.EndParagraph()
		}
		catch (err) {
			oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			oOutput.OutputLn("ERROR: Dokument nije podržan", getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
			oOutput.EndParagraph()
		}

	}
}

function odabirLogo(model) {

	if (getAtrValueFromModel(model, AT_LOGO_KOMPANIJE) == "Omnico")
		var image = Context.createPicture("Omnico.png")
	else if (getAtrValueFromModel(model, AT_LOGO_KOMPANIJE) == "Groser")
		var image = Context.createPicture("Groser.png")
	else if (getAtrValueFromModel(model, AT_LOGO_KOMPANIJE) == "Polimark academy")
		var image = Context.createPicture("Polimark academy.png")
	else if (getAtrValueFromModel(model, AT_LOGO_KOMPANIJE) == "Packsol")
		var image = Context.createPicture("Packsol.png")
	else if (getAtrValueFromModel(model, AT_LOGO_KOMPANIJE) == "HBT")
		var image = Context.createPicture("HBT.png")
	else if (getAtrValueFromModel(model, AT_LOGO_KOMPANIJE) == "Universal brands")
		var image = Context.createPicture("Universal brands.png")
	else if (getAtrValueFromModel(model, AT_LOGO_KOMPANIJE) == "Unisol")
		var image = Context.createPicture("Unisol.png")
	else if (getAtrValueFromModel(model, AT_LOGO_KOMPANIJE) == "Polimark foods")
		var image = Context.createPicture("Polimark foods.png")
	else if (getAtrValueFromModel(model, AT_LOGO_KOMPANIJE) == "PFI")
		var image = Context.createPicture("PFI.png")

	return image
}

function getObjectDefListNamesString(objOccList) {
	var string = " "
	for (var t = 0; t < objOccList.length; t++)
		string = string + objOccList[t].Name(nLocale) + ", "

	string = string.slice(1, -2)

	return string
}