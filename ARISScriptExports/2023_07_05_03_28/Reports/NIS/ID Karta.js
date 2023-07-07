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
var functionObjDefs = ArisData.getSelectedObjDefs()

main(functionObjDefs)

oOutput.WriteReport()

function setupOutputObject(outputObj) {
    outputObj.SetAutoTOCNumbering(true)
    //globalHeader(outputObj)

    //globalFooter(outputObj)
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


function main(functionObjDefs) {
    for (var i = 0; i < functionObjDefs.length; i++) {

        var nazivAtt = functionObjDefs[i].Name(nLocale)
        var oznakaAtt = functionObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("8fa28970-3beb-11ec-7867-00505682218a"), nLocale).getValue()
        //var vlasnikAtt = functionObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("adde7a20-75f9-11e0-61d8-00237d347a8a"), nLocale).getValue()

        var verzijaAtt = functionObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("9f81f3d1-3beb-11ec-7867-00505682218a"), nLocale).getValue()

        var kratakOpisAtt = functionObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("0ce1e020-3be2-11ec-7867-00505682218a"), nLocale).getValue()

        var ciljAtt = functionObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("b312f481-3beb-11ec-7867-00505682218a"), nLocale).getValue()
        var izlazniElementRezAtt = functionObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("34eb5550-3c24-11ec-7867-00505682218a"), nLocale).getValue()
        var procPredhodeAtt = functionObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("c7fb35e1-3bed-11ec-7867-00505682218a"), nLocale).getValue()
        var proSledeAtt = functionObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("d27650e0-3bed-11ec-7867-00505682218a"), nLocale).getValue()
        
        var kljucneAktivnosti = functionObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("4d41eab1-0361-11ed-7867-00505682218a"), nLocale).getValue()

        var nmdLinkAtt = functionObjDefs[i].Attribute(Constants.AT_LINK, nLocale).getValue()
        var kreiraoAtt = functionObjDefs[i].Attribute(Constants.AT_CREATOR, nLocale).getValue()

        var vlasnik

        var resursiIT = functionObjDefs[i].getConnectedObjs([Constants.OT_APPL_SYS_TYPE], Constants.EDGES_INOUT, [Constants.CT_CAN_SUPP_1])
        var resursiITStr = objDefsToString(resursiIT)

        var orgJedinica = functionObjDefs[i].getConnectedObjs([Constants.OT_ORG_UNIT], Constants.EDGES_INOUT, [Constants.CT_IS_TECH_RESP_1])
        var orgJedinicaStr = objDefsToString(orgJedinica)

        var rizici = functionObjDefs[i].getConnectedObjs([Constants.OT_RISK], Constants.EDGES_INOUT, [Constants.CT_OCCUR])
        var riziciStr = objDefsToString(rizici)

        var kljucniPokazatelji = functionObjDefs[i].getConnectedObjs([Constants.OT_KPI], Constants.EDGES_INOUT, [Constants.CT_MEASURED_BY])
        var kljucniPokazateljiStr = objDefsToString(kljucniPokazatelji)

        var standard = functionObjDefs[i].getConnectedObjs([Constants.OT_INFO_CARR], Constants.EDGES_INOUT, [Constants.CT_PROV_INP_FOR])
        var standardStr = objDefsToString(standard)
        
        var ulazi = functionObjDefs[i].getConnectedObjs([Constants.OT_INFO_CARR], Constants.EDGES_IN, null)
        var ulaziFiltered = filterInfoCarr(ulazi)
        var ulaziStr = objDefsToString(ulaziFiltered)
        var izlazi = functionObjDefs[i].getConnectedObjs([Constants.OT_INFO_CARR], Constants.EDGES_OUT, null)
        var izlaziFiltered = filterInfoCarr(izlazi)
        var izlaziStr = objDefsToString(izlaziFiltered)

        var nameParent = functionObjDefs[i].getConnectedObjs([Constants.OT_FUNC], Constants.EDGES_IN, [Constants.CT_IS_PRCS_ORNT_SUPER]) //pripadnost CT_IS_PRCS_ORNT_SUPER
        var nameParentStr = objDefsToString(nameParent)
        
        var pozicija = functionObjDefs[i].getConnectedObjs([Constants.OT_POS], Constants.EDGES_INOUT, [Constants.CT_IS_OWN])
        var pozicijaStr = objDefsToString(pozicija)

        //get assigned FAD
        /* var assignedFADs = functionObjDefs[i].AssignedModels(Constants.MT_FUNC_ALLOC_DGM)
        
        if(assignedFADs.length != 0){ //else nema fada svi ovi ispod prazni neka budu
            var assignedFAD = assignedFADs[0]
            //from FAD get next
            var vlasnik 
            var orgJedinica = assignedFAD.ObjDefListByTypes([Constants.OT_ORG_UNIT]) //objList to string pretvori
            var rizici = assignedFAD.ObjDefListByTypes([Constants.OT_RISK]) //objList to string pretvori
            var kljucniPokazatelji = assignedFAD.ObjDefListByTypes([Constants.OT_KPI]) //objList to string pretvori
            var resursiIT = assignedFAD.ObjDefListByTypes([Constants.OT_KPI]) //objList to string pretvori
            var standard
        } */
        //get connected function, parent 

        //CREATE TABLE
        oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)
        oOutput.TableRow()
        oOutput.ResetFrameStyle()

        oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
        var image = Context.createPicture("logoNIS.png")
        oOutput.OutGraphic(image, -1, 40, 15)
        oOutput.EndParagraph()

        oOutput.TableCell("", 55, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn("ИДЕНТИФИКАЦИОНA КАРТА БИЗНИС ПРОЦЕСА", getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        //oOutput.OutputLn("(Path: " + epcModels[k].Group().Path(nLocale) + ")", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VBOTTOM | Constants.FMT_VTOP | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()

        oOutput.TableCell("", 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn("Верзија:" + verzijaAtt, getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER | Constants.FMT_VCENTER, 0.71)
        //oOutput.OutputLn("(Path: " + epcModels[k].Group().Path(nLocale) + ")", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VBOTTOM | Constants.FMT_VTOP | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()

        oOutput.TableRow()
        oOutput.ResetFrameStyle()
        oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(242, 242, 242), 0, Constants.FMT_VCENTER | Constants.FMT_LEFT, 0)
        oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn("Назив", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VCENTER | Constants.FMT_LEFT | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()
        oOutput.TableCell("", 75, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_LEFT, 0)
        oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn(nazivAtt, getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VCENTER | Constants.FMT_LEFT | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()

        oOutput.TableRow()
        oOutput.ResetFrameStyle()
        oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(242, 242, 242), 0, Constants.FMT_VCENTER | Constants.FMT_LEFT, 0)
        oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn("Ознака", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VCENTER | Constants.FMT_LEFT | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()
        oOutput.TableCell("", 75, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_LEFT | Constants.FMT_VTOP, 0)
        oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn(oznakaAtt, getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VCENTER | Constants.FMT_LEFT | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()

        oOutput.TableRow()
        oOutput.ResetFrameStyle()
        oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(242, 242, 242), 0, Constants.FMT_VCENTER | Constants.FMT_LEFT, 0)
        oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn("Власник", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()
        oOutput.TableCell("", 75, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn(pozicijaStr, getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71) //TODO
        oOutput.EndParagraph()

        oOutput.TableRow()
        oOutput.ResetFrameStyle()
        oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(242, 242, 242), 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn("Организациона јединица", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()
        oOutput.TableCell("", 75, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn(orgJedinicaStr, getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()

        oOutput.TableRow()
        oOutput.ResetFrameStyle()
        oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(242, 242, 242), 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn("Кратак опис", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()
        oOutput.TableCell("", 75, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn(kratakOpisAtt, getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()

        oOutput.TableRow()
        oOutput.ResetFrameStyle()
        oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(242, 242, 242), 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn("Циљ", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()
        oOutput.TableCell("", 75, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn(ciljAtt, getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()

        oOutput.TableRow()
        oOutput.ResetFrameStyle()
        oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(242, 242, 242), 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn("Кључне активности/фазе бизнис процеса", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()
        oOutput.TableCell("", 75, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn(kljucneAktivnosti, getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71) //TODO
        oOutput.EndParagraph()

        oOutput.TableRow()
        oOutput.ResetFrameStyle()
        oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(242, 242, 242), 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn("Pезултат", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()
        oOutput.TableCell("", 75, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn(izlazniElementRezAtt, getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()
        
                oOutput.TableRow()
        oOutput.ResetFrameStyle()
        oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(242, 242, 242), 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn("Кључни улази", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()
        oOutput.TableCell("", 75, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn(ulaziStr, getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()
        
                oOutput.TableRow()
        oOutput.ResetFrameStyle()
        oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(242, 242, 242), 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn("Кључни излали", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()
        oOutput.TableCell("", 75, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn(izlaziStr, getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()

        oOutput.TableRow()
        oOutput.ResetFrameStyle()
        oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(242, 242, 242), 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn("Припадност", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()
        oOutput.TableCell("", 75, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn(nameParentStr, getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()

        oOutput.TableRow()
        oOutput.ResetFrameStyle()
        oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(242, 242, 242), 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn("Процеси који претходе", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()
        oOutput.TableCell("", 75, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn(procPredhodeAtt, getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()

        oOutput.TableRow()
        oOutput.ResetFrameStyle()
        oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(242, 242, 242), 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn("Процеси који следе", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()
        oOutput.TableCell("", 75, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn(proSledeAtt, getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()

        oOutput.TableRow()
        oOutput.ResetFrameStyle()
        oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(242, 242, 242), 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn("Ризици", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()
        oOutput.TableCell("", 75, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn(riziciStr, getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()

        oOutput.TableRow()
        oOutput.ResetFrameStyle()
        oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(242, 242, 242), 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn("Кључни показатељи", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()
        oOutput.TableCell("", 75, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn(kljucniPokazateljiStr, getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()

        oOutput.TableRow()
        oOutput.ResetFrameStyle()
        oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(242, 242, 242), 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn("IT ресурси", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()
        oOutput.TableCell("", 75, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn(resursiITStr, getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()

        oOutput.TableRow()
        oOutput.ResetFrameStyle()
        oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(242, 242, 242), 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn("Стандард", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()
        oOutput.TableCell("", 75, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn(standardStr, getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()

        oOutput.TableRow()
        oOutput.ResetFrameStyle()
        oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(242, 242, 242), 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn("НМД линк", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()
        oOutput.TableCell("", 75, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn(nmdLinkAtt, getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()

        oOutput.TableRow()
        oOutput.ResetFrameStyle()
        oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(242, 242, 242), 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn("Креирао", getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()
        oOutput.TableCell("", 75, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.BeginParagraph(Constants.FMT_LEFT | Constants.FMT_VCENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutputLn(kreiraoAtt, getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0.71)
        oOutput.EndParagraph()

        oOutput.EndTable("", 100, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0);

    }
}

function objDefsToString(objDefs) {
    var strTemp = ""
    for (var i = 0; i < objDefs.length; i++) {
        var nameTemp = objDefs[i].Name(nLocale)
        strTemp = strTemp + nameTemp

        if (i != (objDefs.length - 1)) {
            strTemp = strTemp + ","
        }
    }
    return strTemp
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

function filterInfoCarr(objDefs){
var tempArray = new Array()
for (var i = 0; i < objDefs.length; i++) {
    var defaultSymNum = objDefs[i].getDefaultSymbolNum()  
    if(defaultSymNum == Constants.ST_INFO_CARR){
    tempArray.push(objDefs[i])
    }   
}
return tempArray
}