var oOutput = Context.createOutputObject()
var xlTableActive = false
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

var nLocale = Context.getSelectedLanguage()
var position = ArisData.getSelectedObjDefs()

main(position)

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
	outputObj.TableCell("", 29.11, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
	outputObj.BeginParagraph(Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)

	outputObj.EndParagraph()
	outputObj.TableCell("", 42.62, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_CENTER, 0)
	outputObj.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	outputObj.OutputLn("Job Description", getString("ID_DEFAULT_FONT"), 16, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER | Constants.FMT_VBOTTOM, 0.71)
	outputObj.EndParagraph()
	outputObj.TableCell("", 28.27, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_RIGHT, 0)
	outputObj.BeginParagraph(Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)

	outputObj.OutputLn("", getString("ID_DEFAULT_FONT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT, 0.71)
	outputObj.EndParagraph()
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
	outputObj.TableCell("", 50, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
	outputObj.BeginParagraph(Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
	outputObj.OutputField(Constants.FIELD_DATE, getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT)
	outputObj.OutputLn("", getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0.71)
	outputObj.EndParagraph()
	outputObj.TableCell("", 50, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_RIGHT, 0)
	outputObj.BeginParagraph(Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)
	outputObj.Output(getString("ID_REPORTDEF_2"), getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT, 0.71)
	outputObj.OutputField(Constants.FIELD_PAGE, getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT)
	outputObj.Output(getString("ID_REPORTDEF_3"), getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT, 0.71)
	outputObj.OutputField(Constants.FIELD_NUMPAGES, getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT)
	outputObj.OutputLn("", getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT, 0.71)
	outputObj.EndParagraph()
	if (Context.getSelectedFormat() != Constants.OUTEXCEL)
		outputObj.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
	else
		outputObj.TableRow()
	outputObj.EndFooter()
}

function main() {

	for (var i = 0; i < position.length; i++) {

		var allModelsFromObject = collectModelsFromObject(position[i].OccList())

		var epcModels = new Array();

		for (var j = 0; j < allModelsFromObject.length; j++) {
			if (allModelsFromObject[j].TypeNum() == Constants.MT_EEPC && allModelsFromObject[j].Name(nLocale).substring(0, 4) != "ISO_" && allModelsFromObject[j].Name(nLocale).slice(-4) != "_ISO") {
				epcModels.push(allModelsFromObject[j])
			}
		}

		var orgChartModels = new Array();

		for (var j = 0; j < allModelsFromObject.length; j++) {
			if (allModelsFromObject[j].TypeNum() == Constants.MT_ORG_CHRT) {
				orgChartModels.push(allModelsFromObject[j])
			}
		}

		var roleOcc = new Array()
		for (var j = 0; j < orgChartModels.length; j++) {
			for (var k = 0; k < orgChartModels[j].ObjOccListFilter(Constants.OT_PERS_TYPE).length; k++)
				roleOcc.push(orgChartModels[j].ObjOccListFilter(Constants.OT_PERS_TYPE)[k])
		}

		var allModelsFromRole = new Array()
		for (var j = 0; j < roleOcc.length; j++) {
			var allModelsFromRoleTemp = collectModelsFromObject(roleOcc[j].ObjDef().OccList())

			for (var j1 = 0; j1 < allModelsFromRoleTemp.length; j1++)
				allModelsFromRole.push(allModelsFromRoleTemp[j1])
		}

		if (allModelsFromRole != undefined) {
			for (var j = 0; j < allModelsFromRole.length; j++) {
				if (allModelsFromRole[j].TypeNum() == Constants.MT_EEPC && allModelsFromRole[j].Name(nLocale).substring(0, 4) != "ISO_" && allModelsFromRole[j].Name(nLocale).slice(-4) != "_ISO") {
					epcModels.push(allModelsFromRole[j])
				}
			}
		}

		epcModels = ArisData.Unique(epcModels)

		oOutput.BeginTable(98, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)
		oOutput.TableRow()
		oOutput.ResetFrameStyle()
		oOutput.TableCell("", 100, getString("ID_DEFAULT_FONT"), 12, Constants.C_BLACK, RGB(255, 255, 153), 0, Constants.FMT_VCENTER | Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
		oOutput.BeginParagraph(Constants.FMT_VBOTTOM | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
		oOutput.OutputLn(position[i].Name(nLocale), getString("ID_DEFAULT_FONT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VCENTER | Constants.FMT_CENTER | Constants.FMT_BOLD, 0.71)
		oOutput.EndParagraph()
		oOutput.EndTable("", 100, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
		emptyrow(1)

		for (var k = 0; k < epcModels.length; k++) {

			oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)
			oOutput.TableRow()
			oOutput.ResetFrameStyle()
			oOutput.TableCell("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(204, 255, 204), 0, Constants.FMT_VBOTTOM | Constants.FMT_CENTER, 0)
			oOutput.BeginParagraph(Constants.FMT_VBOTTOM | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
			oOutput.OutputLn(epcModels[k].Name(nLocale), getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VBOTTOM | Constants.FMT_CENTER | Constants.FMT_BOLD, 0.71)
			//oOutput.OutputLn("(Path: " + epcModels[k].Group().Path(nLocale) + ")", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VBOTTOM | Constants.FMT_CENTER | Constants.FMT_BOLD, 0.71)
			oOutput.EndParagraph()
			oOutput.TableRow()
			oOutput.ResetFrameStyle()
			oOutput.TableCell("", 100 / 7, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
			oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
			oOutput.OutputLn("Activity", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VCENTER | Constants.FMT_CENTER | Constants.FMT_BOLD, 0.71)
			oOutput.EndParagraph()
			oOutput.TableCell("", 100 / 7, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
			oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
			oOutput.OutputLn("Participation", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VCENTER | Constants.FMT_CENTER | Constants.FMT_BOLD, 0.71)
			oOutput.EndParagraph()

			/* oOutput.TableCell("", 100 / 8, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
			oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
			oOutput.OutputLn("Additional executors", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VCENTER | Constants.FMT_CENTER | Constants.FMT_BOLD, 0.71)
			oOutput.EndParagraph() */

			oOutput.TableCell("", 100 / 7, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
			oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
			oOutput.OutputLn("Additional participants", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VCENTER | Constants.FMT_CENTER | Constants.FMT_BOLD, 0.71)
			oOutput.EndParagraph()
			oOutput.TableCell("", 100 / 7, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
			oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
			oOutput.OutputLn("Input", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VCENTER | Constants.FMT_CENTER | Constants.FMT_BOLD, 0.71)
			oOutput.EndParagraph()
			oOutput.TableCell("", 100 / 7, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
			oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
			oOutput.OutputLn("Output", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VCENTER | Constants.FMT_CENTER | Constants.FMT_BOLD, 0.71)
			oOutput.EndParagraph()
			oOutput.TableCell("", 100 / 7, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
			oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
			oOutput.OutputLn("Systems", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VCENTER | Constants.FMT_CENTER | Constants.FMT_BOLD, 0.71)
			oOutput.EndParagraph()
			oOutput.TableCell("", 100 / 7, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
			oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
			oOutput.OutputLn("Improvement potential", getString("ID_DEFAULT_FONT"), 7.5, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VCENTER | Constants.FMT_CENTER | Constants.FMT_BOLD, 0.71)
			oOutput.EndParagraph()

			var allFunctionsFromEPC = epcModels[k].ObjOccListFilter(Constants.OT_FUNC)

			var functionsFromEpc = new Array()

			for (var k2 = 0; k2 < allFunctionsFromEPC.length; k2++) {

				var spojeniObjekti = allFunctionsFromEPC[k2].getConnectedObjOccs(new Array(Constants.ST_POS, Constants.ST_POS_1, Constants.ST_POS_2, Constants.ST_EMPL_TYPE, Constants.ST_PERS_TYPE_PIC, Constants.ST_SOLAR_ROLE), Constants.EDGES_INOUT)

				for (var k3 = 0; k3 < spojeniObjekti.length; k3++) {
					if (spojeniObjekti[k3].ObjDef().Name(nLocale) == position[i].Name(nLocale))
						functionsFromEpc.push(allFunctionsFromEPC[k2])

					for (var k4 = 0; k4 < roleOcc.length; k4++) {
						if (roleOcc[k4] != undefined) {
							if (spojeniObjekti[k3].ObjDef().Name(nLocale) == roleOcc[k4].ObjDef().Name(nLocale))
								functionsFromEpc.push(allFunctionsFromEPC[k2])
						}
					}

				}
			}
            
            functionsFromEpc = ArisData.Unique(functionsFromEpc)
            

			for (var k2 = 0; k2 < functionsFromEpc.length; k2++) {

				var additionalExecutors = new Array()
				var additionalParticipants = new Array()
				var spojenePozicijeRole = functionsFromEpc[k2].getConnectedObjOccs(new Array(Constants.ST_POS, Constants.ST_POS_1, Constants.ST_POS_2, Constants.ST_EMPL_TYPE, Constants.ST_PERS_TYPE_PIC, Constants.ST_SOLAR_ROLE), Constants.EDGES_INOUT)

				for (var k3 = 0; k3 < spojenePozicijeRole.length; k3++) {

					if (spojenePozicijeRole[k3].ObjDef().Name(nLocale) == position[i].Name(nLocale)) {

						var imeVeze = connectionName(spojenePozicijeRole[k3])

					}

					for (var k4 = 0; k4 < roleOcc.length; k4++) {
						if (roleOcc[k4] != undefined)
							if (spojenePozicijeRole[k3].ObjDef().Name(nLocale) == roleOcc[k4].ObjDef().Name(nLocale)) {

								var imeVeze = connectionName(spojenePozicijeRole[k3])
								var additionalExecutors = getPositionsFromRole(roleOcc[k4].ObjDef())
							}
					}

					if (spojenePozicijeRole[k3].ObjDef().TypeNum() == Constants.OT_POS && spojenePozicijeRole[k3].ObjDef().Name(nLocale) != position[i].Name(nLocale))
						additionalParticipants.push(spojenePozicijeRole[k3])

					else if (spojenePozicijeRole[k3].ObjDef().TypeNum() == Constants.OT_PERS_TYPE) {
						var zastava = 0
						var addPartTemp = new Array()

						addPartTemp.push(spojenePozicijeRole[k3])

						for (var k4 = 0; k4 < roleOcc.length; k4++) {
							if (roleOcc[k4] != undefined) {
								if (addPartTemp[0].ObjDef().Name(nLocale) == roleOcc[k4].ObjDef().Name(nLocale))
									zastava = 1
							}

						}
						if (zastava != 1)
							additionalParticipants.push(spojenePozicijeRole[k3])

					}
				}

				var inputObjects = functionsFromEpc[k2].getConnectedObjOccs(new Array(Constants.ST_DOC, ArisData.ActiveFilter().UserDefinedSymbolTypeNum(CS_DOCUMENT), Constants.ST_DOCU_PIC, Constants.ST_EMAIL_PIC, Constants.ST_EMAIL_PIC2, Constants.ST_EMAIL_1, Constants.ST_INFO_CARR_EDOC, Constants.ST_INFO_CARR, Constants.ST_NOTE_PIC, Constants.ST_INFO_CARR_NOTE, Constants.ST_PHONE, Constants.ST_PHONE_PIC), Constants.EDGES_IN)

				var outputObjects = functionsFromEpc[k2].getConnectedObjOccs(new Array(Constants.ST_DOC, ArisData.ActiveFilter().UserDefinedSymbolTypeNum(CS_DOCUMENT), Constants.ST_DOCU_PIC, Constants.ST_EMAIL_PIC, Constants.ST_EMAIL_PIC2, Constants.ST_EMAIL_1, Constants.ST_INFO_CARR_EDOC, Constants.ST_INFO_CARR, Constants.ST_NOTE_PIC, Constants.ST_INFO_CARR_NOTE, Constants.ST_PHONE, Constants.ST_PHONE_PIC), Constants.EDGES_OUT)

				var systemsObjects = functionsFromEpc[k2].getConnectedObjOccs(new Array(Constants.ST_APPL_SYS, Constants.ST_APPL_SYS_1, Constants.ST_APPL_SYS_PIC, Constants.ST_APPL_SYS_CLS, Constants.ST_APPL_SYS_TYPE), Constants.EDGES_INOUT)

				oOutput.TableRow()
				oOutput.ResetFrameStyle()
				oOutput.TableCell("", 100 / 7, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
				oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
				oOutput.OutputLn(functionsFromEpc[k2].ObjDef().Name(nLocale), getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71)
				oOutput.EndParagraph()
				oOutput.TableCell("", 100 / 7, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
				oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
				oOutput.OutputLn(imeVeze, getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71)
				oOutput.EndParagraph()

				/*
				oOutput.TableCell("", 100 / 8, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
				oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
				if (additionalExecutors != undefined)
				    for (var z = 0; z < additionalExecutors.length; z++)
				        oOutput.OutputLn(additionalExecutors[z].Name(nLocale), getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71)
				oOutput.EndParagraph() */

				oOutput.TableCell("", 100 / 7, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
				oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
				for (var z = 0; z < additionalParticipants.length; z++)
					oOutput.OutputLn(additionalParticipants[z].ObjDef().Name(nLocale), getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71)
				oOutput.EndParagraph()
				oOutput.TableCell("", 100 / 7, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
				oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
				for (var z = 0; z < inputObjects.length; z++)
					oOutput.OutputLn(inputObjects[z].ObjDef().Name(nLocale), getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71)
				oOutput.EndParagraph()
				oOutput.TableCell("", 100 / 7, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
				oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
				for (var z = 0; z < outputObjects.length; z++)
					oOutput.OutputLn(outputObjects[z].ObjDef().Name(nLocale), getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71)
				oOutput.EndParagraph()
				oOutput.TableCell("", 100 / 7, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
				oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
				for (var z = 0; z < systemsObjects.length; z++)
					oOutput.OutputLn(systemsObjects[z].ObjDef().Name(nLocale), getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71)
				oOutput.EndParagraph()
				oOutput.TableCell("", 100 / 7, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
				oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
				if (functionsFromEpc[k2].ObjDef().AttrList(nLocale, [Constants.AT_IMPROVE])[0] != undefined)
					oOutput.OutputLn(functionsFromEpc[k2].ObjDef().AttrList(nLocale, [Constants.AT_IMPROVE])[0].GetValue(false), getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71)
				oOutput.EndParagraph()
			}
			oOutput.EndTable("", 100, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
			emptyrow(1)
		}
	}
}

function emptyrow(number) {
	for (var i = 0; i < number; i++)
		oOutput.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
}

function connectionName(position) {
	var ime_veze = position.CxnOccList()[0].CxnDef().Attribute(389, nLocale, true).getValue()
	return ime_veze
}

function collectModelsFromObject(p_array) {
	var result = new Array();
	for (var i = 0; i < p_array.length; i++) {
		result.push(p_array[i].Model())
	}
	return ArisData.Unique(result)
}

function getPositionsFromRole(rola) {
	var positionsFromRole = rola.getConnectedObjs([Constants.OT_POS], Constants.EDGES_INOUT, null)
	return positionsFromRole
}

function RGB(r, g, b) {
	return (new java.awt.Color(r / 255.0, g / 255.0, b / 255.0, 1)).getRGB() & 0xFFFFFF
}