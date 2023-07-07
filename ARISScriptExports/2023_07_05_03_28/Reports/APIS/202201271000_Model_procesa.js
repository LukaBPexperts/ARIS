var nLocale = Context.getSelectedLanguage()
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

var startModels = ArisData.getSelectedModels()

main(startModels);

oOutput.WriteReport()

/** Apply default page format settings to output object
 * @param {Output} outputObj The output object
 */

function setupOutputObject(outputObj) {
    outputObj.SetAutoTOCNumbering(true)
    globalHeader(outputObj)

    globalFooter(outputObj)
}

function globalHeader(outputObj) {
    outputObj.BeginHeader()
    if (Context.getSelectedFormat() != Constants.OUTEXCEL)
        //outputObj.BeginTable(100, [40, 30, 30], RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)
        outputObj.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)
    outputObj.TableRow()
    outputObj.ResetFrameStyle()
    outputObj.SetFrameStyle(Constants.FRAME_TOP, 20, Constants.BRDR_NORMAL)
    outputObj.SetFrameStyle(Constants.FRAME_LEFT, 20, Constants.BRDR_NORMAL)
    outputObj.SetFrameStyle(Constants.FRAME_BOTTOM, 20, Constants.BRDR_NORMAL)
    outputObj.SetFrameStyle(Constants.FRAME_RIGHT, 20, Constants.BRDR_NORMAL)

    //outputObj.TableCell("", 3, 1, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_LEFT, 0)
    outputObj.TableCell("", 40, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_LEFT, 0)
    outputObj.BeginParagraph(Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
    var image = Context.createPicture("apisLogo.png")
    outputObj.OutGraphic(image, -1, 55, 20)
    outputObj.EndParagraph()

    //outputObj.TableCell("", 1, 1, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_LEFT, 0)
    outputObj.TableCell("", 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
    outputObj.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    outputObj.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
    outputObj.EndParagraph()

    //outputObj.TableCell("", 1, 1, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_LEFT, 0)
    outputObj.TableCell("", 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
    outputObj.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    outputObj.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
    outputObj.EndParagraph()

    //outputObj.TableCell("", 1, 1, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_LEFT, 0)
    outputObj.TableCell("", 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
    outputObj.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    outputObj.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
    outputObj.EndParagraph()

    //outputObj.TableCell("", 1, 1, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_LEFT, 0)
    outputObj.TableCell("", 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
    outputObj.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    outputObj.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
    outputObj.EndParagraph()

    for (var i = 0; i < 1; i++) {
        outputObj.TableRow()
        //outputObj.TableCell("", 1, 1, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_LEFT, 0)
        outputObj.TableCell("", 40, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
        outputObj.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
        outputObj.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
        outputObj.EndParagraph()

        //outputObj.TableCell("", 1, 1, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_LEFT, 0)
        outputObj.TableCell("", 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
        outputObj.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
        outputObj.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
        outputObj.EndParagraph()

        //outputObj.TableCell("", 1, 1, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_LEFT, 0)
        outputObj.TableCell("", 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
        outputObj.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
        outputObj.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
        outputObj.EndParagraph()

        //outputObj.TableCell("", 1, 1, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_LEFT, 0)
        outputObj.TableCell("", 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
        outputObj.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
        outputObj.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
        outputObj.EndParagraph()

        //outputObj.TableCell("", 1, 1, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_LEFT, 0)
        outputObj.TableCell("", 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
        outputObj.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
        outputObj.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
        outputObj.EndParagraph()
    }


    if (Context.getSelectedFormat() != Constants.OUTEXCEL)
        outputObj.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
    else
        outputObj.TableRow()
    outputObj.EndHeader()
}

function globalFooter(outputObj) {
    outputObj.BeginFooter()
    //if( !isExcel() )
    outputObj.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)
    outputObj.TableRow()
    outputObj.ResetFrameStyle()
    outputObj.SetFrameStyle(Constants.FRAME_TOP, 20, Constants.BRDR_NORMAL)
    outputObj.SetFrameStyle(Constants.FRAME_LEFT, 20, Constants.BRDR_NORMAL)
    outputObj.SetFrameStyle(Constants.FRAME_BOTTOM, 20, Constants.BRDR_NORMAL)
    outputObj.SetFrameStyle(Constants.FRAME_RIGHT, 20, Constants.BRDR_NORMAL)

    outputObj.TableCell("", 50, "Segoe UI", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
    outputObj.BeginParagraph(Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
    outputObj.OutputField(Constants.FIELD_FILENAME, "Segoe UI", 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT)
    outputObj.OutputLn("", "Segoe UI", 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0.71)
    outputObj.EndParagraph()

    outputObj.TableCell("", 50, "Segoe UI", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_RIGHT, 0)
    outputObj.BeginParagraph(Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)
    outputObj.Output("Str. ", "Segoe UI", 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT, 0.71)
    outputObj.OutputField(Constants.FIELD_PAGE, "Segoe UI", 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT)
    outputObj.Output(" od ", "Segoe UI", 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT, 0.71)
    outputObj.OutputField(Constants.FIELD_NUMPAGES, getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT)
    outputObj.OutputLn("", "Segoe UI", 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT, 0.71)
    outputObj.EndParagraph()
    //if( !isExcel() )
    outputObj.EndTable("", 100, "Segoe UI", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
    //else
    //outputObj.TableRow()
    outputObj.EndFooter()
}

function main(startObjDefs) {

    //create first page
    firstPage()

    //create second page
    oOutput.BeginSection(false, Constants.SECTION_DEFAULT)

    oOutput.SetAutoTOCNumbering(true)
    globalHeader(oOutput)
    globalFooter(oOutput)

    emptyrow(2)

    //Izrada/Izmjene:
    oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    oOutput.Output("Izrada/Izmjene:", getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()

    emptyrow(1)

    oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
    oOutput.TableRow()
    oOutput.ResetFrameStyle()
    oOutput.TableCell("", 10, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Verzija", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()
    oOutput.TableCell("", 30, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Opis", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()
    oOutput.TableCell("", 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Ime i Prezime", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()
    oOutput.TableCell("", 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Organizacijska jedinica", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()
    oOutput.TableCell("", 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Datum", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()

    for (var i = 0; i < 5; i++) {
        oOutput.TableRow()
        oOutput.ResetFrameStyle()
        oOutput.TableCell("", 10, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
        oOutput.TableCell("", 30, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
        oOutput.TableCell("", 20, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
        oOutput.TableCell("", 20, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
        oOutput.TableCell("", 20, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
    }
    oOutput.EndTable("", 100, "Segoe UI", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0)

    emptyrow(2)

    //Pregledao/la
    oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    oOutput.Output("Pregledao/la:", getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()

    emptyrow(1)

    oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
    oOutput.TableRow()
    oOutput.ResetFrameStyle()
    oOutput.TableCell("", 10, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Verzija", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Ime i Prezime", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()
    oOutput.TableCell("", 40, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Organizacijska jedinica", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Datum pregleda", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()

    oOutput.TableRow()
    oOutput.ResetFrameStyle()
    oOutput.TableCell("", 10, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
    oOutput.TableCell("", 40, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)

    oOutput.EndTable("", 100, "Segoe UI", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0)

    emptyrow(2)

    //Odobrio/la
    oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    oOutput.Output("Odobrio/la:", getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()

    emptyrow(1)

    oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
    oOutput.TableRow()
    oOutput.ResetFrameStyle()
    oOutput.TableCell("", 10, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Verzija", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Ime i Prezime", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()
    oOutput.TableCell("", 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Potpis", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Funkcija/Organizacijska jedinica", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Datum odobrenja", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()

    oOutput.TableRow()
    oOutput.ResetFrameStyle()
    oOutput.TableCell("", 10, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
    oOutput.TableCell("", 15, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)

    oOutput.EndTable("", 100, "Segoe UI", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0)

    emptyrow(2)

    //Lista za distribuciju:
    //oOutput.addLocalBookmark("TOC");
    oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    oOutput.Output("Lista za distribuciju:", getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()

    emptyrow(1)

    oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
    oOutput.TableRow()
    oOutput.ResetFrameStyle()
    oOutput.TableCell("", 30, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Ime i prezime", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()
    oOutput.TableCell("", 70, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Organizacijska jedinica", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()

    for (var i = 0; i < 3; i++) {
        oOutput.TableRow()
        oOutput.ResetFrameStyle()
        oOutput.TableCell("", 30, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
        oOutput.TableCell("", 75, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
    }
    oOutput.EndTable("", 100, "Segoe UI", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0)

    oOutput.EndSection()

    //TOC
    tableOfContent()

    //**************************************************************************************************
    ////*************************************DOKUMENT***************************************************
    //**************************************************************************************************
    //UVOD
    oOutput.BeginSection(false, Constants.SECTION_DEFAULT)
    globalHeader(oOutput)
    globalFooter(oOutput)

    emptyrow(1)

    //UVOD
    oOutput.BeginParagraphF(getString("HEADING_1"))
    oOutput.Output("1. Uvod", getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()
    
    oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    oOutput.Output("[Napomena – Upisati tekst]", getString("ID_STYLE_RD_DEFAULT"), 11, RGB(237, 28, 36), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_ITALIC, 0)
    oOutput.EndParagraph()
    
    emptyrow(3)
    
    //Preduvjeti
    oOutput.BeginParagraphF(getString("HEADING_2"))
    oOutput.Output("1.1 Preduvjeti", getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()
    
    oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    oOutput.Output("[Napomena – Upisati tekst]", getString("ID_STYLE_RD_DEFAULT"), 11, RGB(237, 28, 36), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_ITALIC, 0)
    oOutput.EndParagraph()
    
    emptyrow(3)
    
    //Napomene
    oOutput.BeginParagraphF(getString("HEADING_2"))
    oOutput.Output("1.2 Napomene", getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()
    
    oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    oOutput.Output("[Napomena – Upisati tekst]", getString("ID_STYLE_RD_DEFAULT"), 11, RGB(237, 28, 36), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_ITALIC, 0)
    oOutput.EndParagraph()
    
    emptyrow(3)

    oOutput.EndSection()
    
    //******************************************************************************************
    //2.3 UC Dijagram
    //******************************************************************************************
    oOutput.BeginSection(false, Constants.SECTION_DEFAULT)
    globalHeader(oOutput)
    globalFooter(oOutput)

    //emptyrow(1)

    oOutput.BeginParagraphF(getString("HEADING_1"))
    oOutput.Output("2 Poslovni procesi", getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()

    //emptyrow(1)

    oOutput.BeginParagraphF(getString("HEADING_2"))
    oOutput.Output("2.1 Elektronička razmjena računa", getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()
    
    var longDescriptionAtt = startModels[0].Attribute(Constants.AT_DESC, nLocale).getValue()
    oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    oOutput.Output(longDescriptionAtt, getString("ID_STYLE_RD_DEFAULT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()
    
    emptyrow(1)
    
    //get first assigned FAD model
    var modelGraphic = startModels[0].Graphic(false, false, nLocale)

    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutGraphic(modelGraphic, -1, 200, 200)
    oOutput.EndParagraph()

    emptyrow(1)

    oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    oOutput.Output("U nastavku slijedi opis koraka krovnog procesa:", getString("ID_STYLE_RD_DEFAULT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()
    
    //var functionObjDefs = startModels[0].ObjDefListByTypes([Constants.OT_FUNC])
    var functionObjOccs = startModels[0].ObjOccListBySymbol([Constants.ST_VAL_ADD_CHN_SML_2, Constants.ST_VAL_ADD_CHN_SML_1])
    
    //sort po koordinatama u modelu
    var sortedFunctionObjOccs = sortByPosition(functionObjOccs)
    
    for(var i=0; i<functionObjOccs.length; i++){
    //var longDescription = startObjDefs[0].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("118b5651-479b-11ec-4a03-10604b88abdc"), nLocale).getHTMLFormattedValue()
    var longDescriptionAtt = functionObjOccs[i].ObjDef().Attribute(Constants.AT_DESC, nLocale).getHTMLFormattedValue()
    oOutput.OutputFormattedText(longDescriptionAtt.getHTML())
    }

    oOutput.EndSection()
    
    //******************************************************************************************
    //2.2 Izdavanje izlaznog računa
    //******************************************************************************************

    oOutput.BeginSection(false, Constants.SECTION_DEFAULT)
    globalHeader(oOutput)
    globalFooter(oOutput)

    emptyrow(1)
    
    for(var i=0; i<functionObjOccs.length; i++){
        
        var bpmnModels = functionObjOccs[i].ObjDef().AssignedModels(Constants.MT_BPMN_COLLABORATION_DIAGRAM)
        
        for(var j=0; j<bpmnModels.length; j++){
            
        oOutput.BeginParagraphF(getString("HEADING_2"))
        oOutput.Output("2." + (i+1) + " " + functionObjOccs[i].ObjDef().Name(nLocale), getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
        oOutput.EndParagraph()

        var longDescriptionAtt = bpmnModels[j].Attribute(Constants.AT_DESC, nLocale).getValue()
        oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
        oOutput.Output(longDescriptionAtt, getString("ID_STYLE_RD_DEFAULT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)
        oOutput.EndParagraph()
            
        //bpmn model picture
        var modelGraphic = bpmnModels[j].Graphic(false, false, nLocale)

        oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
        oOutput.OutGraphic(modelGraphic, -1, 200, 200)
        oOutput.EndParagraph()
        
        emptyrow(1)
        
        //var bpmnObjOccs = bpmnModels[j].ObjOccListBySymbol([Constants.ST_BPMN_TASK, Constants.ST_BPMN_START_EVENT])
        var bpmnObjDefs = bpmnModels[j].ObjDefListByTypes([Constants.OT_FUNC, Constants.OT_EVT])
        bpmnObjDefs = ArisData.sort( bpmnObjDefs, Constants.AT_NAME, Context.getSelectedLanguage())

        for(var k=0; k<bpmnObjDefs.length; k++){
            oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
            oOutput.Output(bpmnObjDefs[k].Name(nLocale), getString("ID_STYLE_RD_DEFAULT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_BOLD, 0)
            oOutput.EndParagraph()
            
            var longDescriptionAtt = bpmnObjDefs[k].Attribute(Constants.AT_DESC, nLocale).getValue()
            oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
            oOutput.Output(longDescriptionAtt, getString("ID_STYLE_RD_DEFAULT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)
            oOutput.EndParagraph()
            
            emptyrow(1)
            
            var bpmnModelsL2 = bpmnObjDefs[k].AssignedModels(Constants.MT_BPMN_COLLABORATION_DIAGRAM)
            
            for(var l=0; l<bpmnModelsL2.length; l++){
                //bpmn model picture
                var modelGraphic = bpmnModelsL2[l].Graphic(false, false, nLocale)

                oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
                oOutput.OutGraphic(modelGraphic, -1, 200, 200)
                oOutput.EndParagraph()
                
                emptyrow(1)

            }    
        }
        
        emptyrow(1)
        
        }  
    }
    
    oOutput.EndSection()
    
    //******************************************************************************************
    //3. Otvorena pitanja
    //******************************************************************************************

    oOutput.BeginSection(false, Constants.SECTION_DEFAULT)
    globalHeader(oOutput)
    globalFooter(oOutput)

    emptyrow(1)
    
    oOutput.BeginParagraphF(getString("HEADING_1"))
    oOutput.Output("3. Otvorena pitanja", getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()
    
    oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    oOutput.Output("[Napomena – Upisati tekst]", getString("ID_STYLE_RD_DEFAULT"), 11, RGB(237, 28, 36), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_ITALIC, 0)
    oOutput.EndParagraph()

    oOutput.EndSection()
}

function RGB(r, g, b) {
    return (new java.awt.Color(r / 255.0, g / 255.0, b / 255.0, 1)).getRGB() & 0xFFFFFF
}

function emptyrow(number) {
    for (var i = 0; i < number; i++)
        oOutput.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
}

function firstPage() {
    oOutput.BeginSection(false, Constants.SECTION_COVER)

    //oOutput.SetAutoTOCNumbering(true)
    globalHeader(oOutput)

    emptyrow(15)

    var sifraProjekta = startModels[0].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("cec36080-6188-11ec-4a03-10604b88abdc"), nLocale).getValue()
    var imeProjekta = startModels[0].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("ddaf4a70-6249-11ec-4a03-10604b88abdc"), nLocale).getValue()
    var modelName = startModels[0].Name(nLocale)

    oOutput.BeginParagraph(Constants.FMT_CENTER, 0, 0, 0, 2, 0)
    oOutput.OutputLn(sifraProjekta + " " + imeProjekta, getString("ID_DEFAULT_FONT"), 21, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0)
    oOutput.OutputLn(modelName, getString("ID_DEFAULT_FONT"), 21, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0)
    oOutput.OutputLn("Model poslovnih procesa", getString("ID_DEFAULT_FONT"), 21, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0)
    oOutput.EndParagraph()

    oOutput.EndSection()
}

function tableOfContent() {

    oOutput.BeginSection(false, Constants.SECTION_INDEX);
    //Header
    // oOutput.BeginHeader();
    // oOutput.BeginTable(100, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_CENTER, 0);
    // oOutput.TableRow();
    // oOutput.TableCell("Content table header", 100, "Segoe UI", 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
    // oOutput.EndTable("Test", 100, "Segoe UI", 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT, 0);
    // oOutput.EndHeader();
    //Footer
    // oOutput.BeginFooter();
    // oOutput.BeginTable(100, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_CENTER, 0);
    // oOutput.TableRow();
    // oOutput.TableCell("Content table footer", 100, "Segoe UI", 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
    // oOutput.EndTable("Test", 100, "Segoe UI", 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT, 0);
    // oOutput.EndFooter();
    //Sets the autonumbering on
    oOutput.SetAutoTOCNumbering(false);
    //Defines the style for the four levels in the TOC
    oOutput.SetTOCFormat(0, "Segoe UI", 12, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_LEFT | Constants.FMT_BOLD, 0, 0, 0, 0);
    oOutput.SetTOCFormat(1, "Segoe UI", 11, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_LEFT | Constants.FMT_BOLD, 5, 5, 2, 2);
    oOutput.SetTOCFormat(2, "Segoe UI", 11, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_LEFT | Constants.FMT_ITALIC, 10, 5, 2, 2);
    oOutput.SetTOCFormat(3, "Segoe UI", 11, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_LEFT, 15, 5, 2, 2);
    //Output text
    emptyrow(1)
    oOutput.OutputLn("SADRŽAJ:", "Segoe UI", 12, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_LEFT | Constants.FMT_BOLD, 0);
    //Adds the table of content
    emptyrow(2)
    oOutput.OutputField(Constants.FIELD_TOC, "Segoe UI", 10, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_LEFT);
    oOutput.EndSection();
}

function sortByPosition(oobjectocclist) {
    for (var i = 0; i < oobjectocclist.length - 1; i++)
        for (var j = i + 1; j < oobjectocclist.length; j++) {
            if (oobjectocclist[i].Y() > oobjectocclist[j].Y()) {
                var temp = oobjectocclist[i]
                oobjectocclist[i] = oobjectocclist[j]
                oobjectocclist[j] = temp
            } else if (oobjectocclist[i].Y() == oobjectocclist[j].Y() && oobjectocclist[i].X() > oobjectocclist[j].X()) {
                var temp = oobjectocclist[i]
                oobjectocclist[i] = oobjectocclist[j]
                oobjectocclist[j] = temp
            }
        }
        return oobjectocclist
}