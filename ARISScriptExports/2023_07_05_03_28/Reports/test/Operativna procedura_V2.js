var g_ooutfile = new __holder(null); // Output object.
var g_nloc
var g_objCurrentModel
var g_blnIsStartEvent = true
var txtRelations = "Relationships"

var modelNames
var orgUnits
var applications
var isSortByOrderNo = true
var activityCount
var vadc
var outputDocumentsNames;

var oOutput = Context.createOutputObject()
var xlTableActive = false
oOutput.DefineF("Information", "Arial", 14, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER | Constants.FMT_VTOP, 0, 0, 1.76, 8.82, 0, 1)
oOutput.DefineF("Title", "Arial", 21, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER | Constants.FMT_VTOP, 0, 0, 1.76, 8.82, 0, 1)
oOutput.DefineF("Normal", "Arial", 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VTOP, 0, 0, 0, 0, 0, 1)
oOutput.DefineF("Heading 4", "Arial", 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_ITALIC | Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_TOCENTRY3, 0, 0, 0, 0, 0, 1)
oOutput.DefineF("Table contents", "Arial", 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VTOP, 0, 0, 0, 0, 0, 1)
oOutput.DefineF("Heading 3", "Arial", 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_ITALIC | Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_TOCENTRY2, 0, 0, 1, 1, 0, 1)
oOutput.DefineF("Header/Footer", "Arial", 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VTOP, 0, 0, 0, 0, 0, 1)
oOutput.DefineF("Heading 2", "Arial", 14, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_ITALIC | Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_TOCENTRY1, 0, 0, 2, 2, 0, 1)
oOutput.DefineF("Heading 1", "Arial", 18, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_TOCENTRY0, 0, 0, 4, 4, 0, 1)
oOutput.DefineF("Table header", "Arial", 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER | Constants.FMT_VTOP, 0, 0, 0, 0, 0, 1)
oOutput.DefineF("REPORT1", "Times New Roman", 14, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0, 0, 0, 0, 0, 0);
oOutput.DefineF("REPORT2", "Times New Roman", 12, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0, 0, 0, 0, 0, 0);
oOutput.DefineF("REPORT3", "Times New Roman", 12, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0, 0, 0, 0, 0, 0);
oOutput.DefineF("REPORT4", "Times New Roman", 11, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0, 0, 0, 0, 0, 0);
setupOutputObject(oOutput)
oOutput.SetTitle(Context.getScriptInfo(Constants.SCRIPT_NAME))

var epcModels

var nLocale = Context.getSelectedLanguage()
showSortOptionsDlg() //#ANTE
//uzima EPC modele i traži sve kontrole i vezane podatke i upisuje u liste
getModelsDetails(ArisData.getSelectedModels());
writeFirstPage(oOutput)
main()

oOutput.WriteReport()

function GetEpcModels(models) {
    for (var i = 0; i < models.length; i++) {
        if (models[i].TypeNum() == Constants.MT_VAL_ADD_CHN_DGM) {
            vadc = models[i]
            var epcModelsTemp = GetModelsFromVACD(models[i])
            GetEpcModels(epcModelsTemp)
        } else if (models[i].TypeNum() == Constants.MT_EEPC) {
            epcModels[epcModels.length] = models[i]
        }
    }

    epcModels = __sort(epcModels, Constants.AT_NAME, Constants.SORT_NONE, Constants.SORT_NONE, nLocale);
}

function GetModelsFromVACD(vacdModel) {
    var objFunctionList = vacdModel.ObjDefListFilter(Constants.OT_FUNC);
    var assignedModelsMain = new Array()
    for (var i = 0; i < objFunctionList.length; i++) {
        var assignedModels = objFunctionList[i].AssignedModels(Constants.MT_EEPC)
        for (var j = 0; j < assignedModels.length; j++) {
            assignedModelsMain[assignedModelsMain.length] = assignedModels[j]
        }
    }

    return assignedModelsMain
}

//Počisti trenutne model i uzme ponovo sve selektirane
function FillModels() {
    epcModels = new Array()
    GetEpcModels(ArisData.getSelectedModels())
}

function getModelsDetails() {
    FillModels()
    if (epcModels.length == 0)
        return

    var epcNamesList = new Array();

    for (var i = 0; i < epcModels.length; i++) {
        AddToListIfNotExists(epcNamesList, epcModels[i].Name(nLocale))
    }

    modelNames = ListToString(epcNamesList);
}

function getObjectsInString(activity, constant, edges, symbolNum) {
    var objTypes = new Array();
    objTypes[0] = constant;
    var objList = new Array();
    var objNames = ""

    var objConnected = activity.getConnectedObjs(objTypes, edges)
    for (var j = 0; j < objConnected.length; j++) {
        if ((symbolNum == null && Constants.OT_APPL_SYS_TYPE == constant) || (symbolNum == null && Constants.ST_INFO_CARR_EDOC != objConnected[j].GetDefaultSymbolNum()) || symbolNum == objConnected[j].GetDefaultSymbolNum()) {
            AddToListIfNotExists(objList, objConnected[j].Attribute(Constants.AT_NAME, nLocale).GetValue(true))
            if (constant == Constants.OT_APPL_SYS_TYPE) {
                var desc = objConnected[j].Attribute(Constants.AT_DESC, nLocale).GetValue(true);
                if (desc.length() > 0)
                    AddToListIfNotExists(objList, "\tŠiri opis: " + desc)
            }
        }
    }

    var counter = 1;
    for (var i = 0; i < objList.length; i++) {
        if (i > 0)
            objNames = objNames + "\n\t";

        if (objList[i].length > 5 && objList[i].substring(1, 5) == "Širi") {
            objNames = objNames + "\t" + objList[i];
        } else {
            objNames = objNames + "\t" + counter.toString() + ". " + objList[i];
            counter = counter + 1
        }
    }

    return objNames
}

function getPositionsInString(activity) {
    var objTypes = new Array();
    objTypes[0] = Constants.OT_POS;
    var activityList = new Array();
    var activities = ""

    var objConnectedApplications = activity.getConnectedObjs(objTypes, Constants.EDGES_INOUTASSIGN)
    for (var j = 0; j < objConnectedApplications.length; j++)
        AddToListIfNotExists(activityList, objConnectedApplications[j].Attribute(Constants.AT_NAME, nLocale).GetValue(true))

    for (var i = 0; i < activityList.length; i++) {
        if (i > 0)
            activities = activities + " i/ili ";
        activities = activities + activityList[i];
    }

    if (activityList.length > 1)
        activities = activities + " izvršavaju aktivnost"; //#ANTE
    else if (activityList.length == 1)
        activities = activities + " izvršava aktivnost"; //#ANTE

    return activities
}

function GetApplicationsFromFunctionList(functionList, applicationsList) {
    var objTypes = new Array();
    objTypes[0] = Constants.OT_APPL_SYS_TYPE;

    for (var i = 0; i < functionList.length; i++) {
        var objConnectedApplications = functionList[i].ObjDef().getConnectedObjs(objTypes, Constants.EDGES_INOUTASSIGN)
        for (var j = 0; j < objConnectedApplications.length; j++)
            AddToListIfNotExists(applicationsList, objConnectedApplications[j].Attribute(Constants.AT_NAME, nLocale).GetValue(true))
    }

    return applicationsList
}

function GetRisksControls(epcModel) {
    var risks = epcModel.ObjOccListFilter(Constants.OT_RISK)
    var objConnectedObjList = new Array()

    for (var i = 0; i < risks.length; i++) {
        var riskDef = risks[i].ObjDef()
        var objTypes = new Array()
        objTypes[0] = Constants.OT_FUNC;
        var conTypes = new Array()
        conTypes[0] = Constants.CT_IS_REDU_BY

        var objConnectedObjects = riskDef.getConnectedObjs(objTypes, Constants.EDGES_OUT, conTypes)
        if (objConnectedObjects.length > 0)
            objConnectedObjList[objConnectedObjList.length] = objConnectedObjects[0]
    }

    return risks
}

function GetControlNameList(riskDef) {
    var objConnectedObjList = new Array()
    var objTypes = new Array()
    objTypes[0] = Constants.OT_FUNC;
    var conTypes = new Array()
    conTypes[0] = Constants.CT_IS_REDU_BY

    var objConnectedObjects = riskDef.getConnectedObjs(objTypes, Constants.EDGES_OUT, conTypes)

    for (var i = 0; i < objConnectedObjects.length; i++) {
        objConnectedObjList[objConnectedObjList.length] = objConnectedObjects[0].Attribute(Constants.AT_NAME, nLocale).GetValue(true)
    }

    return ListToString(objConnectedObjList)
}

function ListToString(stringList) {
    var allInOneString = ""
    for (var i = 0; i < stringList.length; i++) {
        if (i == 0)
            allInOneString = stringList[i]
        else
            allInOneString = allInOneString + "; " + stringList[i]
    }

    if (allInOneString == "")
        return "n/a"
    else
        return allInOneString
}

function ListToStringSeparator(stringList, separator) {
    var allInOneString = ""
    for (var i = 0; i < stringList.length; i++) {
        if (i == 0)
            allInOneString = stringList[i]
        else
            allInOneString = allInOneString + separator + " " + stringList[i]
    }

    if (allInOneString == "")
        return "n/a"
    else
        return allInOneString
}

function AddToListIfNotExists(stringList, str) {
    var max = stringList.length
    var alreadyInTheList = false;

    if (max == 0)
        stringList[stringList.length] = str
    else {
        for (var i = 0; i < max; i++) {
            if (stringList[i] == str) {
                alreadyInTheList = true;
                break;
            }
        }

        if (alreadyInTheList == false)
            stringList[stringList.length] = str
    }
}

/** Apply default page format settings to output object
 * @param {Output} outputObj The output object
 */
function setupOutputObject(outputObj) {
    outputObj.SetPageWidth(210.1)
    outputObj.SetPageHeight(297.2)
    outputObj.SetLeftMargin(20)
    outputObj.SetRightMargin(20)
    outputObj.SetTopMargin(20)
    outputObj.SetBottomMargin(20)
    outputObj.SetDistHeader(10)
    outputObj.SetDistFooter(10)
    outputObj.SetAutoTOCNumbering(true)
    globalHeader(outputObj)

    globalFooter(outputObj)
}

function globalHeader(outputObj) {
    outputObj.BeginHeader()
    if (Context.getSelectedFormat() != Constants.OUTEXCEL)
        outputObj.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)
    outputObj.TableRow()
    outputObj.ResetFrameStyle()
    outputObj.TableCell("", 50, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
    outputObj.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    outputObj.OutputLn("Operativna procedura", getString("ID_DEFAULT_FONT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
    outputObj.EndParagraph()
    outputObj.TableCell("", 50, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
    outputObj.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    var image = Context.createPicture("BIB logo.jpg")
    outputObj.OutGraphic(image, -1, 59, 10)

    outputObj.EndParagraph()
    if (Context.getSelectedFormat() != Constants.OUTEXCEL)
        outputObj.EndTable("", 0, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
    else
        outputObj.TableRow()
    outputObj.EndHeader()
}

function globalFooter(outputObj) {

    var time = GetCurrentTime();
    outputObj.BeginFooter()
    if (Context.getSelectedFormat() != Constants.OUTEXCEL)
        outputObj.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)
    outputObj.TableRow()
    outputObj.ResetFrameStyle()
    outputObj.TableCell("", 22.25, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_CENTER, 0)
    outputObj.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    outputObj.OutputField(Constants.FIELD_DATE, "Arial", 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER)
    outputObj.OutputLn("", "Arial", 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
    outputObj.EndParagraph()
    outputObj.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    outputObj.OutputLn(time, "Arial", 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
    outputObj.EndParagraph()
    outputObj.TableCell("", 55.51, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
    outputObj.TableCell("", 22.25, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
    outputObj.BeginParagraph(Constants.FMT_CENTER, 1, 1, 0, 0, 0)
    outputObj.Output("Str. ", "Arial", 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 1)
    outputObj.OutputField(Constants.FIELD_PAGE, "Arial", 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER)
    outputObj.Output(" od ", "Arial", 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 1)
    outputObj.OutputField(Constants.FIELD_NUMPAGES, "Arial", 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER)
    outputObj.OutputLn("", "Arial", 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 1)
    outputObj.EndParagraph()
    if (Context.getSelectedFormat() != Constants.OUTEXCEL)
        outputObj.EndTable("", 100, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
    else
        outputObj.TableRow()
    outputObj.EndFooter()
}

//get the Item path
function getItemPath(searchItem, intItemType) {

    if (intItemType == Constants.CID_OBJDEF) {
        return searchItem.Group().Path(nLocale);
    }
    if (intItemType == Constants.CID_MODEL) {
        return searchItem.Group().Path(nLocale);
    }
    if (intItemType == Constants.CID_GROUP) {
        return searchItem.Path(nLocale);
    }
}

function writeEmptyRows(numberOfRows) {
    var i = 0;

    for (i = 0; i < numberOfRows; i++) {
        oOutput.BeginParagraphF("Normal")
        oOutput.OutputLnF("", "Normal")
        oOutput.EndParagraph()
    }
}

function GetCurrentTime() {
    var currentTime = new Date()
    var hours = currentTime.getHours()
    var minutes = currentTime.getMinutes()

    var suffix = "AM";
    if (hours >= 12) {
        suffix = "PM";
        hours = hours - 12;
    }

    if (hours == 0)
        hours = 12;

    if (minutes < 10)
        minutes = "0" + minutes

    return hours.toString() + ":" + minutes.toString() + " " + suffix;
}

function GetCurrentDate() {
    var d = new Date();
    var curr_date = d.getDate();
    var curr_month = d.getMonth() + 1;
    var curr_year = d.getFullYear();
    return curr_date.toString() + "/" + curr_month.toString() + "/" + curr_year.toString();
}

/** writes a cover page to output object
 * @param {Output} outputObj The ARIS output object
 */
function writeFirstPage(outputObj) {
    if (epcModels.length == 0)
        return

    var model = epcModels[0];
    var Item_Name
    var Item_Type
    var Item_Path
    if (vadc == null) {
        Item_Name = model.Name(nLocale);
        Item_Type = model.KindNum();
        Item_Path = getItemPath(model, Item_Type);
    } else {
        Item_Name = vadc.Name(nLocale);
        Item_Type = vadc.KindNum();
        Item_Path = getItemPath(vadc, Item_Type);
    }


    writeEmptyRows(9);
    outputObj.BeginParagraph(Constants.FMT_CENTER, 0, 0, 0, 0, 0)
    outputObj.OutputLn("OPERATIVNA INSTRUKCIJA", "Verdana", 28, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER | Constants.FMT_BOLD, 0) //#ANTE
    outputObj.EndParagraph()

    writeEmptyRows(2);
    outputObj.BeginParagraph(Constants.FMT_CENTER, 0, 0, 0, 0, 0)
    outputObj.OutputLn(Item_Name, "Verdana", 24, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
    outputObj.EndParagraph()

    writeEmptyRows(9);

    outputObj.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)
    outputObj.TableRow()
    outputObj.ResetFrameStyle()
    outputObj.TableCell("Struktura", 20, "Verdana", 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_BOLD, 0)
    outputObj.TableCell(Item_Path, 80, "Verdana", 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0);
    outputObj.TableRow()
    outputObj.ResetFrameStyle()
    outputObj.TableCell("Datum", 20, "Verdana", 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_BOLD, 0)
    outputObj.TableCell(GetCurrentDate(), 80, "Verdana", 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
    outputObj.EndTable("", 100, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0)

    pageBreak(outputObj);
}

function pageBreak(outputObj) {
    outputObj.OutputField(Constants.FIELD_NEWPAGE, "Arial", 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT)
}

function RGB(r, g, b) {
    return (new java.awt.Color(r / 255.0, g / 255.0, b / 255.0, 1)).getRGB() & 0xFFFFFF
}

function writeTextBoldItalic(text, size) {
    oOutput.OutputLn(text, "Arial", size, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_ITALIC | Constants.FMT_LEFT, 0)
}

function writeText(text, size) {
    oOutput.OutputLn(text, "Arial", size, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)
}

function writeTextUnderline(text, size) {
    oOutput.OutputLn(text, "Arial", size, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_UNDERLINE | Constants.FMT_LEFT, 0)
}

function writeTextSameRow(text, text2, size, stil) {
    oOutput.BeginParagraph(Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
    if (stil == "NB") {
        oOutput.Output(text, "Arial", size, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0.71)
        oOutput.Output(text2, "Arial", size, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0.71)
    } else if (stil == "NU") {
        oOutput.Output(text, "Arial", size, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0.71)
        oOutput.Output(text2, "Arial", size, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_UNDERLINE | Constants.FMT_LEFT, 0.71)
    }
    oOutput.EndParagraph()
}

function writeFunctionData(oFunctionDef) {
    //kreso
    activityCount = activityCount + 1
    writeEmptyRows(1);
    writeTextBoldItalic("- Aktivnost " + activityCount.toString() + ".", 10); //#ANTE
    var positions = getPositionsInString(oFunctionDef)
    writeText(positions, 10);
    var activityName = oFunctionDef.Attribute(Constants.AT_NAME, g_nloc).GetValue(true);

    activityName = removeNewLines(activityName)
    writeTextBoldItalic(activityName, 10);
    if (oFunctionDef.getDefaultSymbolNum() == getString("ATTIVIA_DI_CONTROLO_NUMBER"))
        writeTextBoldItalic("[Kontrola]", 10);

    if (oFunctionDef.Attribute(Constants.AT_DESC, g_nloc).GetValue(true) != "") { //#ANTE ako je opis aktivnosti prazan nemoj ga ispisati
        writeEmptyRows(1);
        writeTextBoldItalic("- Širi opis aktivnosti:", 10);
        writeText(oFunctionDef.Attribute(Constants.AT_DESC, g_nloc).GetValue(true), 10);
    }
    
    var appAndTrans = getObjectsInString(oFunctionDef, Constants.OT_APPL_SYS_TYPE, Constants.EDGES_INOUTASSIGN, null)
    if (appAndTrans.length > 0) {
        writeEmptyRows(1);
        writeTextBoldItalic("\t- Korišćene aplikacije:", 10); //#ANTE
        writeText("\t" + appAndTrans, 10);
    }

    var ulazneForme = getObjectsInString(oFunctionDef, Constants.OT_INFO_CARR, Constants.EDGES_IN, null)
    if (ulazneForme.length > 0) {
        writeEmptyRows(1);
        writeTextBoldItalic("\t- Ulazne forme:", 10);
        writeText("\t" + ulazneForme, 10);
    }

    var electrDoc = getObjectsInString(oFunctionDef, Constants.OT_INFO_CARR, Constants.EDGES_INOUT, Constants.ST_INFO_CARR_EDOC)
    if (electrDoc.length > 0) {
        writeEmptyRows(1);
        writeTextBoldItalic("\t- Application documents used:", 10);
        writeText("\t" + electrDoc, 10);
    }

    getOutputDocuments(oFunctionDef);
    var izlazniDokumenti = getObjectsInString(oFunctionDef, Constants.OT_INFO_CARR, Constants.EDGES_OUT, null)
    if (izlazniDokumenti.length > 0) {
        writeEmptyRows(1);
        writeTextBoldItalic("\t- Izlazni dokumenti:", 10);
        writeText("\t" + izlazniDokumenti, 10);
    }

    writeCases(oFunctionDef)
}

function writeCases(oFunctionDef) {
    var objTypes = new Array();
    objTypes[0] = Constants.OT_EVT;

    var connectedEvents = oFunctionDef.getConnectedObjs(objTypes, Constants.EDGES_OUT);
    if (connectedEvents.length == 0) {
        objTypes[0] = Constants.OT_RULE;
        var rules = oFunctionDef.getConnectedObjs(objTypes, Constants.EDGES_OUT);
        if (rules.length > 0) {
            objTypes[0] = Constants.OT_EVT;
            var connectedEvents = rules[0].getConnectedObjs(objTypes, Constants.EDGES_OUT);
        }
    }

    for (var index = 0; index < connectedEvents.length; index++) { //#ANTE dodan flag, izbacen ispis COIDCE ispisuje se ime funkcije, ako je zadnji event i iza njega nema funkcija onda ispisi njega
        var flag = 0
        var event = connectedEvents[index];
        // writeEmptyRows(1)
        // writeTextSameRow("U slučaju sledećega ",event.Attribute(Constants.AT_NAME,nLocale).GetValue(true),10,"NB") //#ANTE
        var objTypes2 = new Array();
        objTypes2[0] = Constants.OT_FUNC;
        var connectedFunctions = event.getConnectedObjs(objTypes2, Constants.EDGES_OUT);
        if (connectedFunctions.length > 0) {
            var funkcija = connectedFunctions[0];
            var codice = funkcija.Attribute(Constants.AT_NAME, nLocale).GetValue(true) //#ANTE izbacen atribut njihov "CODICE", stavljeno ime funkcije
                // if(codice.length()>0)
            if (funkcija != undefined) {
                writeEmptyRows(1)
                writeTextSameRow("U slučaju sledećega ", event.Attribute(Constants.AT_NAME, nLocale).GetValue(true), 10, "NB") //#ANTE
                writeTextSameRow("Proces se nastavlja sledećom aktivnosti ", codice, 10, "NU") //#ANTE
                flag = 1;
            }

        }
        if (flag == 0) {
            writeEmptyRows(1)
            writeTextSameRow("Završni događaj za ovaj tok je: ", event.Attribute(Constants.AT_NAME, nLocale).GetValue(true), 10, "NB") //#ANTE
        }
    }
}

function getOutputDocuments(oFunctionDef) {
    var objTypes = new Array();
    objTypes[0] = Constants.OT_INFO_CARR;

    var objConnectedOutputDocsList = oFunctionDef.getConnectedObjs(objTypes, Constants.EDGES_OUT);
    for (var index = 0; index < objConnectedOutputDocsList.length; index++)
        AddToListIfNotExists(outputDocumentsNames, "- " + objConnectedOutputDocsList[index].Name(nLocale))
}

function removeNewLines(text) {
    var sChar1 = String.fromCharCode(10);
    var sChar2 = String.fromCharCode(13);
    text = text.replace(sChar1, "").replace(sChar2, " ");
    return text;
}

function getActivatesEvent(p_Object) {
    var objTypes = new Array();
    objTypes[0] = Constants.OT_EVT;
    var conTypes = new Array();
    conTypes[0] = Constants.CT_ACTIV_1;
    conTypes[1] = Constants.CT_ACTIV_2;
    conTypes[2] = Constants.CT_ACTIV_3;
    conTypes[3] = Constants.CT_ACTIVATES;

    var objConnectedObjectsTemp = p_Object.getConnectedObjs(objTypes, Constants.EDGES_IN, conTypes);
    var objConnectedObjects = new Array();
    for (var i = 0; i < objConnectedObjectsTemp.length; i++) {
        objConnectedObjects[objConnectedObjects.length] = objConnectedObjectsTemp[i];
    }

    if (objConnectedObjects.length > 0) {
        var objEvent = objConnectedObjects[0];

        if (isEventPositionAbove(objEvent, p_Object))
            return objEvent;
        else
            return null;
    } else
        return null;
}

function isEventPositionAbove(objEvent, objObject) {
    var objModels = new Array();
    objModels[0] = g_objCurrentModel;

    objEventOccList = objEvent.OccList(objModels);
    objObjectOcc = objObject.OccList(objModels)[0];

    for (var i = 0; i < objEventOccList.length; i++) {
        if (objEventOccList[i].Y() < objObjectOcc.Y())
            return true;
    }

    return false;
}

function IsStartEvent(objPrecedingEvent) {
    if (objPrecedingEvent == null) {
        g_blnIsStartEvent = false
        return
    }

    var ocxnocclist = objPrecedingEvent.CxnList(Constants.EDGES_OUT);
    for (var i = 0; i < ocxnocclist.length; i++) {
        if (ocxnocclist[i].Attribute(Constants.AT_UA_VAL_10, nLocale).GetValue(true) == "Iniziale") {
            g_blnIsStartEvent = true
            return
        }
    }
    g_blnIsStartEvent = false
}

function main() {
    g_nloc = nLocale;
    g_ooutfile.value = oOutput;

    if (epcModels.length > 0) {
        for (var i = 0; i < epcModels.length; i++) {
            activityCount = 0;
            if (i > 0)
                pageBreak(oOutput)
                
                //Slika modela
            oOutput.BeginParagraph(Constants.FMT_CENTER, 0, 0, 0, 2, 0)
            oOutput.OutGraphic(epcModels[i].Graphic(false, false, Context.getSelectedLanguage()), 50, 220, 220) //#ANTE
            oOutput.EndParagraph()

            //outputmodeldata(epcModels[i]);  //#ANTE izbaceni podaci koji su se ispisivali ispod slike modela

            //kreso
            writeEmptyRows(1);
            oOutput.Output("OPERATIVNA INSTRUKCIJA", "Arial", 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_ITALIC | Constants.FMT_LEFT, 0) //#ANTE
            writeEmptyRows(1);
            writeTextBoldItalic("Tok aktivnost započinje sledećim događajem:", 10); //#ANTE
            var initialeEvent = getInitialeEventName(epcModels[i]);
            writeTextBoldItalic("- " + initialeEvent, 10);

            //Get functions
            var functionsInModel = GetFunctionsInModel(epcModels[i])
            if (isSortByOrderNo)
                functionsInModel = sortByOrderNo(functionsInModel)
            else
                sortByPosition(functionsInModel)

            if (functionsInModel.length > 0) {
                outputDocumentsNames = new Array()
                for (var j = 0; j < functionsInModel.length; j++) {
                    writeFunctionData(functionsInModel[j].ObjDef())
                    if (j < functionsInModel.length - 1)
                        writeText("______________________________________________________________________________________", 10);
                }

                var outputDocumentsNamesInSting = ListToStringSeparator(outputDocumentsNames, "\n\t");
                writeTextBoldItalic("_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _", 10);
                if (outputDocumentsNamesInSting.length > 0) {
                    writeEmptyRows(1)
                    writeTextBoldItalic("IZLAZNI DOKUMENTI", 10);
                    writeText("\t" + outputDocumentsNamesInSting, 10);
                    writeEmptyRows(1);
                }

                writeRisks(epcModels[i]);
            }
        }
    }
}

function writeRisks(epcModel) {
    var objConnectedObj = epcModel.ObjOccListFilter(Constants.OT_RISK) //GetRisksControls(epcModel)

    if (objConnectedObj.length > 0) {
        writeEmptyRows(1)
        writeTextBoldItalic("LISTA RIZIKA", 10);
        for (var index = 0; index < objConnectedObj.length; index++) {
            writeTextBoldItalic("\t- " + objConnectedObj[index].ObjDef().Attribute(Constants.AT_NAME, nLocale).GetValue(true), 10)
            var desc = objConnectedObj[index].ObjDef().Attribute(Constants.AT_DESC, nLocale).GetValue(true);
            if (desc.length() > 0)
                writeText("\t" + desc, 10)
            writeEmptyRows(1)
        }
    }
}

function getInitialeEventName(model) {
    var objTypes = new Array();
    objTypes[0] = Constants.OT_EVT;

    var eventsNamesList = new Array()
    var objEventsList = model.ObjDefListFilter(objTypes);

    if (objEventsList.length > 0) {
        for (var index = 0; index < objEventsList.length; index++) {
            var ocxnocclist = objEventsList[index].CxnList(Constants.EDGES_OUT);
            for (var i = 0; i < ocxnocclist.length; i++) { 
                if (ocxnocclist[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum(getString("AT_TIPOLOGIA CONNETTORE")), nLocale).GetValue(false) == "Iniziale") //#ANTE promijenio atribut AT_UA_VAL_10 u custom atribut Tipologia Connettore
                {
                    eventsNamesList[eventsNamesList.length] = objEventsList[index].Attribute(Constants.AT_NAME, nLocale).GetValue(true);
                    continue
                }
            }
        }
    }

    var eventsInString = ListToStringSeparator(eventsNamesList, ",");

    return eventsInString;
}

function GetFunctionsInModel(ocurrentmodel) {
    var objFunctionList = ocurrentmodel.ObjOccListFilter(Constants.OT_FUNC);
    var functions = new Array()
    for (var i = 0; i < objFunctionList.length; i++) {
        var isNotInterface = true
        if (objFunctionList[i].SymbolNum() == Constants.ST_PRCS_IF)
            isNotInterface = false

        var isNotSystemFunc = true
        if (objFunctionList[i].SymbolNum() == Constants.ST_SYS_FUNC_ACT)
            isNotSystemFunc = false

        var isControlFunction = false
        if (objFunctionList[i].SymbolNum() == getString("ATTIVIA_DI_CONTROLO_NUMBER"))
            isControlFunction = true

        if (isNotInterface)
            functions[functions.length] = objFunctionList[i]
    }

    return functions
}

//Slika modela #ANTE ova funkcija je zakomentirana u main-u.
function outputmodeldata(ocurrentmodel) {
    g_ooutfile.value.OutputLn("", "Times New Roman", 8, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
    g_ooutfile.value.OutputLn(ocurrentmodel.Name(g_nloc), "Times New Roman", 12, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0);
    g_ooutfile.value.OutputLn("", "Times New Roman", 8, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0);

    // Graphic of the model.
    //graphicout(g_ooutfile, ocurrentmodel); //#ANTE ovo je ispisivalo ruznu sliku, izbačeno i u main funkciju dodan novi ispis slike

    g_ooutfile.value.OutputLn("", "Times New Roman", 8, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
    g_ooutfile.value.OutputLnF("Figure: " + ocurrentmodel.Name(g_nloc), "REPORT4");
    g_ooutfile.value.OutputLn("", "Times New Roman", 8, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
}


function sortByOrderNo(oobjectocclist) {
    oobjectocclist = ArisData.sort(oobjectocclist, ArisData.ActiveFilter().UserDefinedAttributeTypeNum(getString("AT_ORDER_NO")), nLocale) //#ANTE dodan atriubt order no. u funkcije na modelu i napravljena nova funkcija za sortiranje
    //oobjectocclist = __sort(oobjectocclist, Constants.AT_USER_ATTR5, Constants.SORT_NONE, Constants.SORT_NONE, nLocale);  //#ANTE ovo je original funkcija koja se nalazila tu ali nije radila
    return oobjectocclist
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
}

function objectHasNoOutputConnections(objOcc) {
    return (objOcc.OutEdges(Constants.EDGES_ALL).length == 0);
}

function objectIsConnectedToEndProcessInterface(objOcc) {
    var SymbolTypes = new Array();
    SymbolTypes[0] = Constants.ST_PRCS_IF;

    var objProcIfOccList = objOcc.getConnectedObjOccs(SymbolTypes, Constants.EDGES_OUT);

    if (objProcIfOccList.length != 0) {
        var objProcIfOcc = objProcIfOccList[0];

        if (objProcIfOcc.OutEdges(Constants.EDGES_ALL).length == 0)
            return true;
    }

    return false;
}

function objectIsConnectedToEndOperator(objOcc) {
    var SymbolTypes = new Array();
    SymbolTypes[0] = Constants.ST_OPR_AND_1;
    SymbolTypes[1] = Constants.ST_OPR_OR_1;
    SymbolTypes[2] = Constants.ST_OPR_XOR_1;

    var objOperatorOccList = objOcc.getConnectedObjOccs(SymbolTypes, Constants.EDGES_OUT);

    if (objOperatorOccList.length != 0) {
        var objOperatorOcc = getLastOperator(objOperatorOccList[0]);

        if (objectIsConnectedToEndProcessInterface(objOperatorOcc))
            return true;
    }

    return false;
}

function getLastOperator(objOcc) {
    var SymbolTypes = new Array();
    SymbolTypes[0] = Constants.ST_OPR_AND_1;
    SymbolTypes[1] = Constants.ST_OPR_OR_1;
    SymbolTypes[2] = Constants.ST_OPR_XOR_1;

    var objOperatorOccList = objOcc.getConnectedObjOccs(SymbolTypes, Constants.EDGES_OUT);

    if (objOperatorOccList.length != 0)
        return getLastOperator(objOperatorOccList[0]);
    else
        return objOcc;
}


function getEndEventTxtOperator(objEndEvent, objModel) {
    var objModels = new Array();
    objModels[0] = objModel;
    var txtOperator = " and ";
    var objEventOcc = objEndEvent.OccList(objModels)[0];
    var connectedOperatorBellowOcc = getOperatorBellow(objEventOcc);

    if (connectedOperatorBellowOcc != null) {
        if (connectedOperatorBellowOcc.SymbolNum() == Constants.ST_OPR_AND_1)
            txtOperator = " and ";
        else {
            if (connectedOperatorBellowOcc.SymbolNum() == Constants.ST_OPR_OR_1 || connectedOperatorBellowOcc.SymbolNum() == Constants.ST_OPR_XOR_1)
                txtOperator = " or ";
        }
    }

    return txtOperator;
}


function getOperatorBellow(objEventOcc) {
    var SymbolTypes = new Array();
    SymbolTypes[0] = Constants.ST_OPR_AND_1;
    SymbolTypes[1] = Constants.ST_OPR_OR_1;
    SymbolTypes[2] = Constants.ST_OPR_XOR_1;

    var objConnectedObjectsOcc = objEventOcc.getConnectedObjOccs(SymbolTypes, Constants.EDGES_OUT);

    if (objConnectedObjectsOcc.length > 0)
        return objConnectedObjectsOcc[0];
    else
        return null;
}

function getActivatesOperator(p_Object) {
    var objTypes = new Array();
    objTypes[0] = Constants.OT_RULE;
    var conTypes = new Array();
    conTypes[0] = Constants.CT_ACTIV_1;
    conTypes[1] = Constants.CT_ACTIV_2;
    conTypes[2] = Constants.CT_ACTIV_3;
    conTypes[3] = Constants.CT_ACTIVATES;

    var objConnectedObjects = p_Object.getConnectedObjs(objTypes, Constants.EDGES_IN, conTypes);

    if (objConnectedObjects.length > 0)
        return objConnectedObjects[0];
    else
        return null;
}

function getLinksOperator(p_Object) {
    var objTypes = new Array();
    objTypes[0] = Constants.OT_RULE;
    var conTypes = new Array();
    conTypes[0] = Constants.CT_LNK_1;
    conTypes[1] = Constants.CT_LNK_2;
    conTypes[2] = Constants.CT_LNK_3;
    conTypes[3] = Constants.CT_LNK_4;

    var objConnectedObjects = p_Object.getConnectedObjs(objTypes, Constants.EDGES_IN, conTypes);

    if (objConnectedObjects.length > 0)
        return objConnectedObjects[0];
    else
        return null;
}

function showSortOptionsDlg() {
    var nuserdlg = 0; // Variable for the user dialog box

    var userdialog = Dialogs.createNewDialogTemplate(0, 0, 280, 150, "Function sorting option");
    userdialog.GroupBox(10, 10, 260, 75, "Please select the sorting method:");
    userdialog.OptionGroup("options1");
    userdialog.OptionButton(20, 30, 200, 15, "Sort by order number");
    userdialog.OptionButton(20, 50, 200, 15, "Sort topologicaly(Y and X axis)");
    userdialog.OKButton(50, 110, 75, 21);

    var dlg = Dialogs.createUserDialog(userdialog);
    dlg.setDlgValue("options1", 0);

    nuserdlg = __toInteger(Dialogs.show(__currentDialog = dlg));
    // Showing dialog and waiting for confirmation with OK
    var izabrano = __toInteger(dlg.getDlgValue("options1"));

    if (izabrano == 0)
        isSortByOrderNo = true
    else
        isSortByOrderNo = false
}