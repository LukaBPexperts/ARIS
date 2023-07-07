var nLocale = Context.getSelectedLanguage()
var outputFile = Context.createOutputObject()
var ads = Context.getComponent("ADS")
var filter = ArisData.ActiveFilter()
var selectedDatabases = ArisData.getSelectedDatabases()
var selectedModels = selectedDatabases[0].Find(Constants.SEARCH_MODEL, null, -1, nLocale, "", Constants.SEARCH_CMP_EQUAL)
var modelJsonOut = new Array()
var modelJsonOutTemp = new Array()
var dirPath = "C:\\ArisJSON\\"
var docPath = "C:\\ArisJSON\\documents\\"
var picPath = "C:\\ArisJSON\\images\\"
createDir(dirPath)
createDir(docPath)
createDir(picPath)

extractPictures(selectedModels)

var docModels = getModelsWithAttDocs(selectedModels)
extractDocFromAtt("model", docModels, docPath)

var docObjects = getObjectsWithAttDocs(selectedModels)
extractDocFromAtt("object", docObjects, docPath)

var adsLinks = getADSLinks(selectedModels)
extractDocFromAds("ads", adsLinks, docPath)

for (var i = 0; i < selectedModels.length; i++) {
    Context.writeStatus("Creating model.json   Model: " + (i + 1) + '/' + (selectedModels.length + 1))

    var modelType = selectedModels[i].Type()

    var modelTypeApiName = filter.getAPIName(Constants.CID_MODEL, selectedModels[i].TypeNum())

    //var modelName = selectedModels[i].Name(nLocale)
    var modelGUID = selectedModels[i].GUID()

    var attributes = arrayToJsonArray(getAttJsonString(selectedModels[i]))

    var objectsOcc = selectedModels[i].ObjOccListFilter()

    var elementLinksArray = new Array()

    var textOccs = selectedModels[i].TextOccList()
    var textOccsArray = new Array()
    for (var j = 0; j < textOccs.length; j++) {
        var content = getContentLang(textOccs[j])
        var xPos = textOccs[j].X()
        var yPos = textOccs[j].Y()

        textOccsArray.push(jSonTextLabel(content, xPos, yPos))

        content = null
        xPos = null
        yPos = null

    }
    var textLabels = arrayToJsonArray(textOccsArray)
    textOccsArray = null;

    var gfxObjects = selectedModels[i].getGfxObjects()
    var gfxObjectsArray = new Array()

    for (var j = 0; j < gfxObjects.length; j++) {

        var effect3D = gfxObjects[j].get3DEffect()
        var brushColor = gfxObjects[j].getBrushColor()
        var fillColor = gfxObjects[j].getFillColor()
        var height = gfxObjects[j].getHeight()
        var width = gfxObjects[j].getWidth()
        var xPos = gfxObjects[j].getX()
        var yPos = gfxObjects[j].getY()
        var zOrder = gfxObjects[j].getZOrder()
        var typeNum = gfxObjects[j].getTypeNum()

        if (typeNum == 0) {
            var type = "RoundedRectangle"
            var roundnessX = gfxObjects[j].getRoundnessX()
            var roundnessY = gfxObjects[j].getRoundnessY()
            var shaded = gfxObjects[j].isShaded()
            gfxObjectsArray.push(jSonGfxRoundedRectangleObjects(effect3D, brushColor, fillColor, height, width, xPos, yPos, zOrder, type, roundnessX, roundnessY, shaded))
        }

        if (typeNum == 1) {
            var type = "Polygon"
            var points = gfxObjects[j].getPoints()
            var pathArray = getPathJsonString(points)
            var filled = gfxObjects[j].isFilled()
            gfxObjectsArray.push(jSonGfxPolygonObjects(effect3D, brushColor, fillColor, height, width, xPos, yPos, zOrder, type, pathArray, filled))
        }
    }

    var gfxObjectsJsonArray = arrayToJsonArray(gfxObjectsArray)
    gfxObjectsArray = null;

    var oleObjects = selectedModels[i].ComObjOccs()
    var imageObjectsArray = new Array()

    for (var j = 0; j < oleObjects.length; j++) {
        var imageGUID = getGUIDFromString(oleObjects[j].ObjectID())
        var imageHeight = oleObjects[j].Height()
        var imageWidth = oleObjects[j].Width()
        var imageX = oleObjects[j].X()
        var imageY = oleObjects[j].Y()
        var imageZOrder = oleObjects[j].getZOrder()

        imageObjectsArray.push(jSonPictures(imageGUID, imageHeight, imageWidth, imageX, imageY, imageZOrder))

    }
    var imageJsonArray = arrayToJsonArray(imageObjectsArray)
    imageObjectsArray = null;

    for (var j = 0; j < objectsOcc.length; j++) {
        var id = getGUIDFromString(objectsOcc[j].ObjectID())
        var definitionId = objectsOcc[j].ObjDef().GUID()
        var symbolType = objectsOcc[j].SymbolName()
        var symbolTypeApiName = filter.getAPIName(Constants.CID_OBJOCC, objectsOcc[j].SymbolNum())

        var xPos = objectsOcc[j].X();
        var yPos = objectsOcc[j].Y();
        var heightPos = objectsOcc[j].Height();
        var widthPos = objectsOcc[j].Width();
        var zIndex = objectsOcc[j].getZOrder();

        var attPlacement = objectsOcc[j].AttrOccList()
        var attPlacementArray = jSonAttPlacement(attPlacement)
        var attPlacementJsonArray = arrayToJsonArray(attPlacementArray)

        objectsOcc[j] = null;

        var position = jSonObjectPosition(xPos, yPos, widthPos, heightPos);
        xPos = null;
        yPos = null;
        widthPos = null;
        heightPos = null;

        var objekt = jSonObject(id, definitionId, symbolType, position, zIndex, attPlacementJsonArray, symbolTypeApiName)
        id = null;
        definitionId = null;
        symbolType = null;
        position = null;
        zIndex = null;

        elementLinksArray.push(objekt)
        objekt = null;
    }

    objectsOcc = null;
    var cxnOcc = selectedModels[i].CxnOccList()

    selectedModels[i] = null;

    for (var j = 0; j < cxnOcc.length; j++) {
        var id = getGUIDFromString(cxnOcc[j].ObjectID())
        var definitionId = cxnOcc[j].getDefinition().GUID()
        var cxnType = cxnOcc[j].getDefinition().ActiveType()
        var cxnTypeApiName = filter.getAPIName(Constants.CID_CXNDEF, cxnOcc[j].getDefinition().TypeNum())

        var objectFrom = getGUIDFromString(cxnOcc[j].getSource().ObjectID())
        var objectTo = getGUIDFromString(cxnOcc[j].getTarget().ObjectID())

        var points = cxnOcc[j].getPoints()
        var path = getPathJsonString(points)

        var visible = cxnOcc[j].getVisible()

        var attPlacement = cxnOcc[j].AttrOccList()
        var attPlacementArray = jSonAttPlacement(attPlacement)
        var attPlacementJsonArray = arrayToJsonArray(attPlacementArray)

        cxnOcc[j] = null;

        var objekt = jSonConnection(id, definitionId, cxnType, objectFrom, objectTo, path, visible, attPlacementJsonArray, cxnTypeApiName)
        id = null;
        definitionId = null;
        cxnType = null;
        objectFrom = null;
        objectTo = null;
        path = null;
        visible = null;

        elementLinksArray.push(objekt)
        objekt = null;
    }
    cxnOcc = null;

    var elementLinksJsonArray = arrayToJsonArray(elementLinksArray)
    elementLinksArray = null;

    var modelJson = jSonModel(modelType, elementLinksJsonArray, modelGUID, attributes, modelTypeApiName, textLabels, gfxObjectsJsonArray, imageJsonArray)
    modelType = null;
    elementLinksJsonArray = null;
    modelGUID = null;
    attributes = null;
    textLabel = null;

    modelJsonOutTemp.push(modelJson)
    modelJson = null;

}

writeUsingFileWriter(modelJsonOutTemp, dirPath, "model.json")
modelJsonOutTemp = null;
outputFile.OutputTxt(dirPath + "\model.json" + " --> SUCCESSFULY CREATED ON ARIS SERVER" + "\n");

var elementLinksArray = new Array()
var objectsDef = new Array()
var cxnDef = new Array()
var selectedModels = selectedDatabases[0].Find(Constants.SEARCH_MODEL, null, -1, nLocale, "", Constants.SEARCH_CMP_EQUAL)

for (var i = 0; i < selectedModels.length; i++) {
    Context.writeStatus("Creating definition.json   Get objects and connections    Model: " + (i + 1) + '/' + (selectedModels.length + 1))

    var objectsOcc = selectedModels[i].ObjOccListFilter()
    for (var j = 0; j < objectsOcc.length; j++) {
        objectsDef.push(objectsOcc[j].ObjDef())
        objectsOcc[j] = null;
    }
    objectsOcc = null;

    var cxnOcc = selectedModels[i].CxnOccList()
    for (var j = 0; j < cxnOcc.length; j++) {
        cxnDef.push(cxnOcc[j].getDefinition())
        cxnOcc[j] = null;
    }
    cxnOcc = null;

    selectedModels[i] = null;
}

var objectsDef = ArisData.Unique(objectsDef)
var cxnDef = ArisData.Unique(cxnDef)

for (var j = 0; j < objectsDef.length; j++) {
    Context.writeStatus("Creating definition.json   Get data from objects    Object: " + (j + 1) + '/' + (objectsDef.length + 1))
    var uuid = objectsDef[j].GUID()
    var objectType = objectsDef[j].Type()

    var objectTypeApiName = filter.getAPIName(Constants.CID_OBJDEF, objectsDef[j].TypeNum())

    var attributes = arrayToJsonArray(getAttJsonString(objectsDef[j]))

    attList = null;
    var assigmentsList = objectsDef[j].AssignedModels()
    objectsDef[j] = null;
    var assigments = arrayToJsonArray(getAssigmentsJsonString(assigmentsList))
    assigmentsList = null;

    var objekt = jSonObjCxnDef(uuid, objectType, attributes, assigments, objectTypeApiName);
    uuid = null;
    objectType = null;
    attributes = null;
    assigments = null;
    objectTypeApiName = null;
    elementLinksArray.push(objekt)
    objekt = null;
}

for (var j = 0; j < cxnDef.length; j++) {
    Context.writeStatus("Creating definition.json   Get data from connections   Connection: " + (j + 1) + '/' + (cxnDef.length + 1))
    var uuid = cxnDef[j].GUID()
    var cxnType = cxnDef[j].ActiveType()
    var cxnTypeApiName = filter.getAPIName(Constants.CID_CXNDEF, cxnDef[j].TypeNum())

    var attributes = arrayToJsonArray(getAttJsonString(cxnDef[j]))

    var assigmentsList = cxnDef[j].AssignedModels()
    cxnDef[j] = null;
    var assigments = arrayToJsonArray(getAssigmentsJsonString(assigmentsList))
    assigmentsList = null;

    var objekt = jSonObjCxnDef(uuid, cxnType, attributes, assigments, cxnTypeApiName);
    uuid = null;
    objectType = null;
    attributes = null;
    assigments = null;
    cxnTypeApiName = null;
    writeUsingFileWriter
    elementLinksArray.push(objekt)
    objekt = null;
}

writeUsingFileWriter(elementLinksArray, dirPath, "definition.json")
elementLinksArray = null;
outputFile.OutputTxt(dirPath + "\definition.json" + " --> SUCCESSFULY CREATED ON ARIS SERVER" + "\n");

//var groups = new Array()
var explorerItemJsonArray = new Array()
var selectedModels = selectedDatabases[0].Find(Constants.SEARCH_MODEL, null, -1, nLocale, "", Constants.SEARCH_CMP_EQUAL)

for (var i = 0; i < selectedModels.length; i++) {
    Context.writeStatus("Creating explorer.json   Get data from models   Model: " + (i + 1) + '/' + (selectedModels.length + 1))

    /*    var groupModel = selectedModels[i].Group()

        while (groupModel.Name(nLocale) != "") {
            groups.push(groupModel)
            groupModel = groupModel.Parent()
        } */

    var databaseName = selectedModels[i].Group().Database()
    var itemType = "model"
    var modelType = selectedModels[i].Type();
    var itemTypeApiName = filter.getAPIName(Constants.CID_MODEL, selectedModels[i].TypeNum())
    var attributes = null

    var nameModel = selectedModels[i].Name(nLocale)
    var parentId = selectedModels[i].Group().GUID()
    var uuid = selectedModels[i].GUID();

    var deleted = false
    var creationDate = selectedModels[i].Attribute(Constants.AT_CREAT_TIME_STMP, nLocale).getValue()
    var createdBy = selectedModels[i].Attribute(Constants.AT_CREATOR, nLocale).getValue()
    var modificationDate = selectedModels[i].Attribute(Constants.AT_LAST_CHNG_2, nLocale).getValue()
    var modifiedBy = selectedModels[i].Attribute(Constants.AT_LUSER, nLocale).getValue()

    selectedModels[i] = null;

    var meta = jSonExplorerItemMeta(deleted, creationDate, createdBy, modificationDate, modifiedBy)
    deleted = null;
    creationDate = null;
    createdBy = null;
    modificationDate = null;
    modifiedBy = null;

    var explorerItemJson = jSonExplorerItem(databaseName, itemType, modelType, nameModel, attributes, parentId, uuid, meta, itemTypeApiName)
    databaseName = null;
    itemType = null;
    modelType = null;
    nameModel = null;
    parentId = null;
    uuid = null;
    meta = null;
    itemTypeApiName = null;

    explorerItemJsonArray.push(explorerItemJson)
    explorerItemJson = null;
}

selectedModels = null;

var allObjectsDef = selectedDatabases[0].Find(Constants.SEARCH_OBJDEF, null, -1, nLocale, "", Constants.SEARCH_CMP_EQUAL)
for (var i = 0; i < allObjectsDef.length; i++) {
    Context.writeStatus("Creating explorer.json   Get data from objects   Object: " + (i + 1) + '/' + (allObjectsDef.length + 1))

    /*  var groupModel = allObjectsDef[i].Group()

        while (groupModel.Name(nLocale) != "") {
            groups.push(groupModel)
            groupModel = groupModel.Parent()
        } */

    var databaseName = allObjectsDef[i].Database()
    var itemType = "elementDefinition"
    var modelType = "";
    var nameObject = allObjectsDef[i].Name(nLocale)
    var parentId = allObjectsDef[i].Group().GUID()
    var uuid = allObjectsDef[i].GUID();
    var itemTypeApiName = filter.getAPIName(Constants.CID_OBJDEF, allObjectsDef[i].TypeNum())
    var attributes = null

    var deleted = false
    var creationDate = allObjectsDef[i].Attribute(Constants.AT_CREAT_TIME_STMP, nLocale).getValue()
    var createdBy = allObjectsDef[i].Attribute(Constants.AT_CREATOR, nLocale).getValue()
    var modificationDate = allObjectsDef[i].Attribute(Constants.AT_LAST_CHNG_2, nLocale).getValue()
    var modifiedBy = allObjectsDef[i].Attribute(Constants.AT_LUSER, nLocale).getValue()

    var meta = jSonExplorerItemMeta(deleted, creationDate, createdBy, modificationDate, modifiedBy)
    deleted = null;
    creationDate = null;
    createdBy = null;
    modificationDate = null;
    modifiedBy = null;

    var explorerItemJson = jSonExplorerItem(databaseName, itemType, modelType, nameObject, attributes, parentId, uuid, meta, itemTypeApiName)
    databaseName = null;
    itemType = null;
    modelType = null;
    nameObject = null;
    parentId = null;
    uuid = null;
    meta = null;
    itemTypeApiName = null;

    explorerItemJsonArray.push(explorerItemJson)
    explorerItemJson = null;
    allObjectsDef[i] = null;
}

var allGroups = selectedDatabases[0].Find(Constants.SEARCH_GROUP)

//var groupsDef = ArisData.Unique(groups)
var groupsDef = ArisData.Unique(allGroups)

for (var i = 0; i < groupsDef.length; i++) {
    Context.writeStatus("Creating explorer.json   Get data from groups   Group: " + (i + 1) + '/' + (groupsDef.length + 1))

    var databaseName = groupsDef[i].Database()
    var itemType = "folder"
    var modelType = "";
    var nameGroup = groupsDef[i].Name(nLocale)
    var attributes = arrayToJsonArray(getAttJsonString(groupsDef[i]))

    var parentId = groupsDef[i].Parent().GUID()
    var uuid = groupsDef[i].GUID();
    var itemTypeApiName = "GROUP"

    var deleted = false
    var creationDate = groupsDef[i].Attribute(Constants.AT_CREAT_TIME_STMP, nLocale).getValue()
    var createdBy = groupsDef[i].Attribute(Constants.AT_CREATOR, nLocale).getValue()
    var modificationDate = groupsDef[i].Attribute(Constants.AT_LAST_CHNG_2, nLocale).getValue()
    var modifiedBy = groupsDef[i].Attribute(Constants.AT_LUSER, nLocale).getValue()

    var meta = jSonExplorerItemMeta(deleted, creationDate, createdBy, modificationDate, modifiedBy)
    deleted = null;
    creationDate = null;
    createdBy = null;
    modificationDate = null;
    modifiedBy = null;

    var explorerItemJson = jSonExplorerItem(databaseName, itemType, modelType, nameGroup, attributes, parentId, uuid, meta, itemTypeApiName)
    databaseName = null;
    itemType = null;
    modelType = null;
    nameGroup = null;
    parentId = null;
    uuid = null;
    meta = null;
    attributes = null;

    explorerItemJsonArray.push(explorerItemJson)
    explorerItemJson = null;
    groupsDef[i] = null;
}


writeUsingFileWriter(explorerItemJsonArray, dirPath, "explorer.json")
outputFile.OutputTxt(dirPath + "\explorer.json" + " --> SUCCESSFULY CREATED ON ARIS SERVER" + "\n");

var zipPath = "C:\\"
zipDirectory(zipPath, dirPath, "testZip.zip")


outputFile.WriteReport();

function zipDirectory(pathName, rootPath, zipName) {
  var zipFile = new java.io.File(pathName + zipName);
  var zipOut = new java.util.zip.ZipOutputStream(new java.io.FileOutputStream(zipFile));
  var rootDir = new java.io.File(rootPath);
  var buffer = java.lang.reflect.Array.newInstance(java.lang.Byte.TYPE, 1024);

  function addZipEntry(file, parentPath) {
    if (file.isFile()) {
      var inputStream = new java.io.FileInputStream(file);
      var zipEntry = new java.util.zip.ZipEntry(parentPath + file.getName());
      zipOut.putNextEntry(zipEntry);
      var len;
      while ((len = inputStream.read(buffer)) > 0) {
        zipOut.write(buffer, 0, len);
      }
      inputStream.close();
      zipOut.closeEntry();
    } else if (file.isDirectory()) {
      var files = file.listFiles();
      for (var i = 0; i < files.length; i++) {
        addZipEntry(files[i], parentPath + file.getName() + "/");
      }
    }
  }

  addZipEntry(rootDir, "");

  zipOut.close();
}
function getPathJsonString(points) {
    var pathArray = new Array()
    for (var i = 0; i < points.length; i++) {

        var x = points[i].getX()
        var y = points[i].getY()
        pathArray.push(jSonConnectionPoint(x, y))
    }

    return arrayToJsonArray(pathArray)
}

function getAttJsonString(selection) {
    //selection = model, objectDef, group
    var attributes = new Array()
    var attList = new Array()
    var values = ""

    var allLocales = getLanguageList()

    for (var i = 0; i < allLocales.length; i++) {
        var tempAttList = selection.AttrList(allLocales[i], null)
        var attList = attList.concat(tempAttList)
    }

    var attList = ArisData.Unique(attList)

    var valuesArray = new org.json.JSONArray();

    for (var i = 0; i < attList.length; i++) {
        var valuesArray = new org.json.JSONArray();

        var value = attList[i].getValue()
        var language = attList[i].Locale()
        var languageName = getLocaleName(language)
        var attName = attList[i].Type()
        var order = i + 1;
        var type = attList[i].TypeNum().toString()
        var typeApiName = filter.getAPIName(Constants.CID_ATTRDEF, attList[i].TypeNum())

        valuesArray = jSonObjCxnAttValuesLang(language, languageName, value, valuesArray);

        if (attList[i].LanguageDependence() != 1) {
            for (var j = 0; j < attList.length; j++) {
                if (attList[i].Type() == attList[j].Type() && attList[i].Locale() != attList[j].Locale()) {
                    var value = attList[j].getValue()
                    var language = attList[j].Locale()
                    var languageName = getLocaleName(language)
                    valuesArray = jSonObjCxnAttValuesLang(language, languageName, value, valuesArray);
                    attList.splice(j, 1)
                    j--
                }
            }
        }
        attributes.push(jSonObjCxnAttributes(attName, order.toString(), type, valuesArray.toString(2), typeApiName))
    }
    return attributes;
}

function getAssigmentsJsonString(assigmentsList) {
    var assigments = new Array()
    for (var i = 0; i < assigmentsList.length; i++) {
        var id = assigmentsList[i].GUID()
        assigments.push(jSonAssigments(id))
    }
    return assigments;
}


function getGUIDFromString(aString) {
    var result = java.util.UUID.nameUUIDFromBytes(aString.getBytes()).toString();
    return result;
}

function jSonObjectPosition(left, top, width, height) {
    var obj = new org.json.JSONObject();
    obj.put("left", left);
    obj.put("top", top);
    obj.put("width", width);
    obj.put("height", height);

    return obj.toString(2);
}

function jSonAttPlacement(attPlacement) {
    var attPlacementArray = new Array()

    for (var i = 0; i < attPlacement.length; i++) {

        var attPlacTypeNum = attPlacement[i].AttrTypeNum()
        var attPlacOrderNum = attPlacement[i].getOrderNum()

        var attPlacRotation = attPlacement[i].getRotation()
        var attPlacTextBoxWidth = attPlacement[i].getTextBoxSize().getWidth()
        var attPlacTextBoxHeight = attPlacement[i].getTextBoxSize().getHeight()
        var attPlacTextSizeWidth = attPlacement[i].getTextSize(nLocale).getWidth()
        var attPlacTextSizeHeight = attPlacement[i].getTextSize(nLocale).getHeight()

        var attPlacPort = attPlacement[i].GetPortOptions()[0]
        var portTypeApiName = getPortTypeApiName(attPlacPort);

        var attPlacView = attPlacement[i].GetPortOptions()[1]
        var viewTypeApiName = getViewTypeApiName(attPlacView);

        var attPlacOffsetX = attPlacement[i].OffsetX()
        var attPlacOffsetY = attPlacement[i].OffsetY()

        var obj = new org.json.JSONObject();
        obj.put("typeNum", attPlacTypeNum);
        obj.put("orderNum", attPlacOrderNum);
        obj.put("rotation", attPlacRotation);
        obj.put("textBoxWidth", attPlacTextBoxWidth);
        obj.put("textBoxHeight", attPlacTextBoxHeight);
        obj.put("textSizeWidth", attPlacTextSizeWidth);
        obj.put("textSizeHeight", attPlacTextSizeHeight);

        obj.put("port", portTypeApiName);
        obj.put("view", viewTypeApiName);

        obj.put("offsetX", attPlacOffsetX);
        obj.put("offsetY", attPlacOffsetY);

        attPlacementArray.push(obj.toString(2));
        obj = null;
    }

    return attPlacementArray
}

function getViewTypeApiName(attPlacView) {
    var apiName = ""

    if (attPlacView == 1) {
        var apiName = "ATTROCC_SYMBOL"
    } else if (attPlacView == 2) {
        var apiName = "ATTROCC_NAME)"
    } else if (attPlacView == 0) {
        var apiName = "ATTROCC_TEXT"
    } else if (attPlacView == 4) {
        var apiName = "ATTROCC_FORBID_CHARFORMATS"
    }
    return apiName;
}

function getPortTypeApiName(attPlacPort) {
    var apiName = ""

    if (attPlacPort == 1) {
        var apiName = "ATTROCC_TOPLEFT"
    } else if (attPlacPort == 2) {
        var apiName = "ATTROCC_TOP"
    } else if (attPlacPort == 3) {
        var apiName = "ATTROCC_TOPRIGHT"
    } else if (attPlacPort == 4) {
        var apiName = "ATTROCC_LEFT"
    } else if (attPlacPort == 5) {
        var apiName = "ATTROCC_CENTER"
    } else if (attPlacPort == 6) {
        var apiName = "ATTROCC_RIGHT"
    } else if (attPlacPort == 7) {
        var apiName = "ATTROCC_BOTTOMLEFT"
    } else if (attPlacPort == 8) {
        var apiName = "ATTROCC_BOTTOM"
    } else if (attPlacPort == 9) {
        var apiName = "ATTROCC_BOTTOMRIGHT"
    } else if (attPlacPort == 10) {
        var apiName = "ATTROCC_CENTERTOP"
    } else if (attPlacPort == 11) {
        var apiName = "ATTROCC_CENTERBOTTOM"
    } else if (attPlacPort == 12) {
        var apiName = "ATTROCC_PORT_FREE"
    }
    return apiName
}

function jSonObject(id, definitionId, symbolType, position, zIndex, attPlacementJsonArray, symbolTypeApiName) {

    var obj = new org.json.JSONObject();
    var positionObj = new org.json.JSONObject(position);
    var attPlacementObj = new org.json.JSONArray(attPlacementJsonArray);

    obj.put("id", id);
    obj.put("definitionId", definitionId);
    obj.put("symbolType", symbolType);
    obj.put("position", positionObj);
    obj.put("zIndex", zIndex);
    obj.put("attPlacement", attPlacementObj);
    obj.put("symbolTypeApiName", symbolTypeApiName);

    return obj.toString(2);
}

function jSonConnectionPoint(xX, yY) {

    var obj = new org.json.JSONObject();
    obj.put("x", xX);
    obj.put("y", yY);

    return obj.toString(2);
}

function arrayToJsonArray(elementLinks) {

    var array = new org.json.JSONArray();
    for (var i = 0; i < elementLinks.length; i++) {
        var obj = new org.json.JSONObject(elementLinks[i]);
        array.put(obj);
        obj = null;
        elementLinks[i] = null;
    }
    return array.toString(2);
}

function jSonConnection(id, definitionId, symbolType, objectFrom, objectTo, path, visible, attPlacementJsonArray, cxnTypeApiName) {
    var obj = new org.json.JSONObject();
    var pathObj = new org.json.JSONArray(path);
    var attPlacementObj = new org.json.JSONArray(attPlacementJsonArray);

    obj.put("id", id);
    obj.put("definitionId", definitionId);
    obj.put("symbolType", symbolType);
    obj.put("objectFrom", objectFrom);
    obj.put("objectTo", objectTo);
    obj.put("path", pathObj);
    obj.put("visible", visible);
    obj.put("symbolTypeApiName", cxnTypeApiName);
    obj.put("attPlacement", attPlacementObj);

    return obj.toString(2);
}

function jSonModel(type, elementLinks, uuid, attributes, modelTypeApiName, textLabel, gfxObjectsJsonArray, imageJson) {
    var obj = new org.json.JSONObject();

    var textLabelObj = new org.json.JSONArray(textLabel);
    var gfxObjectsObj = new org.json.JSONArray(gfxObjectsJsonArray);
    var imageJsonArray = new org.json.JSONArray(imageJson);
    var elementLinksArray = new org.json.JSONArray(elementLinks);
    var attributesArray = new org.json.JSONArray(attributes);

    obj.put("type", type);
    obj.put("elementLinks", elementLinksArray);
    obj.put("uuid", uuid);
    obj.put("attributes", attributesArray);
    obj.put("typeApiName", modelTypeApiName);
    obj.put("textLabels", textLabelObj);
    obj.put("gfxObjects", gfxObjectsObj);
    obj.put("images", imageJsonArray);

    return obj.toString(2);
}


function writeUsingFileWriter(data, pathName, fileName) {
    var file = new java.io.File(pathName + fileName);
    var fr = new java.io.OutputStreamWriter(new java.io.FileOutputStream(file), java.nio.charset.StandardCharsets.UTF_8);

    // fstream = new FileWriter(mergedFile, false);
    // fstream = new OutputStreamWriter(new FileOutputStream(mergedFile), StandardCharsets.UTF_8);

    fr.write("[\n");
    for (var i = 0; i < data.length; i++) {
        Context.writeStatus("Writing file. Element " + i + " od " + data.length)
        fr.write(data[i]);
        if (i != (data.length - 1)) {
            fr.write(",\n")
        }
    }
    fr.write("\n]");
    fr.close();
}

function createDir(path) {
    var directory = new java.io.File(path);
    if (!directory.exists()) {
        directory.mkdir();
    }

    var files = directory.listFiles();

    for (var i = 0; i < files.length; i++) {
        files[i].delete()
    }

}

function jSonObjCxnAttValues(language, value) {

    var obj = new org.json.JSONObject();
    var valuesArray = new org.json.JSONArray();

    obj.put("language", language);
    obj.put("value", value);
    valuesArray.put(obj);

    return valuesArray.toString(2);
}

function jSonObjCxnAttValuesLang(language, languageName, value, valuesArray) {

    var obj = new org.json.JSONObject();

    obj.put("language", language);
    obj.put("languageName", languageName);
    obj.put("value", value);
    valuesArray.put(obj);

    return valuesArray;
}

function jSonObjCxnAttributes(attName, order, type, values, typeApiName) {

    var obj = new org.json.JSONObject();
    var valuesArray = new org.json.JSONArray(values);

    obj.put("name", attName);
    obj.put("order", order);
    obj.put("type", type);
    obj.put("values", valuesArray);
    obj.put("typeApiName", typeApiName);

    return obj.toString(2);
}

function jSonAssigments(id) {

    var obj = new org.json.JSONObject();
    obj.put("id", id);

    return obj.toString(2);
}

function jSonObjCxnDef(uuid, objectType, attributes, assignments, objectTypeApiName) {

    var obj = new org.json.JSONObject();
    var attributesArray = new org.json.JSONArray(attributes);
    var assigmentsArray = new org.json.JSONArray(assignments);

    obj.put("uuid", uuid);
    obj.put("objectType", objectType);
    obj.put("attributes", attributesArray);
    obj.put("assignments", assigmentsArray);
    obj.put("objectTypeApiName", objectTypeApiName);

    return obj.toString(2);
}

function jSonExplorerItemMeta(deleted, creationDate, createdBy, modificationDate, modifiedBy) {

    var obj = new org.json.JSONObject();

    obj.put("deleted", deleted);
    obj.put("creationDate", creationDate);
    obj.put("createdBy", createdBy);
    obj.put("modificationDate", modificationDate);
    obj.put("modifiedBy", modifiedBy);

    return obj.toString(2);
}

function jSonExplorerItem(databaseName, itemType, modelType, modelName, attributes, parentId, uuid, meta, typeApiName) {

    var obj = new org.json.JSONObject();
    var metaObj = new org.json.JSONObject(meta);
    if (attributes != null)
        var attributesArray = new org.json.JSONArray(attributes);

    obj.put("databaseName", databaseName);
    obj.put("itemType", itemType);
    obj.put("modelType", modelType);
    obj.put("name", modelName);
    if (attributes != null)
        obj.put("attributes", attributesArray);
    obj.put("parentId", parentId);
    obj.put("uuid", uuid);
    obj.put("meta", metaObj);
    obj.put("itemTypeApiName", typeApiName);

    return obj.toString(2);
}

function extractDocFromAtt(type, docContext, path) {
    for (var i = 0; i < docContext.length; i++) {
        Context.writeStatus("Extracting document from " + type + ": " + (i + 1) + '/' + (docContext.length + 1))

        var docAtt = docContext[i].Attribute(Constants.AT_DOCUMENTATION, nLocale)
        var extDoc = docAtt.getExternalDocument()
        var extDocFileExtension = extDoc.getFileExtension()
        var byteData = extDoc.getDocument()
        var docName = docContext[i].Name(nLocale) + "." + extDocFileExtension
        var zipName = type + "_" + docContext[i].GUID() + ".zip"

        docContext[i] = null;

        zipDocFile(byteData, path, docName, zipName)
    }
    docContext = null;
}

function extractDocFromAds(type, adsLinks, docAdsPath) {
    for (var i = 0; i < adsLinks.length; i++) {
        Context.writeStatus("Extracting document from " + type + ": " + (i + 1) + '/' + (adsLinks.length + 1))

        var adsDoc = ads.getDocumentByHyperlink(adsLinks[i])
        adsLinks[i] = null;

        if (adsDoc != null) {
            var adsDocContent = adsDoc.getDocumentContent()
            var bytes = new org.apache.commons.io.IOUtils.toByteArray(adsDocContent);
            var zipName = type + "_" + adsDoc.getIdentifier() + ".zip"
            var docMetaInfo = adsDoc.getDocumentMetaInfo()
            var docName = docMetaInfo.getFileName()

            zipDocFile(bytes, docAdsPath, docName, zipName)
        }
    }
}

function zipDocFile(byteData, pathName, docName, zipName) {
    var zipFile = new java.io.File(pathName + zipName);
    var zipOut = new java.util.zip.ZipOutputStream(new java.io.FileOutputStream(zipFile));
    var zipEntry = new java.util.zip.ZipEntry(docName);
    zipOut.putNextEntry(zipEntry);
    zipOut.write(byteData, 0, byteData.length);
    zipOut.closeEntry();
    zipOut.close();
}

function getModelsWithAttDocs(selectedModels) {
    var modelsTemp = new Array()

    for (var i = 0; i < selectedModels.length; i++) {
        if (selectedModels[i].Attribute(Constants.AT_DOCUMENTATION, nLocale).getExternalDocument().isValid() == true) {
            modelsTemp.push(selectedModels[i])
        }
    }
    var modelsTemp = ArisData.Unique(modelsTemp)
    return modelsTemp
}

function getObjectsWithAttDocs(selectedModels) {
    var objectsDefTemp = new Array()

    for (var i = 0; i < selectedModels.length; i++) {

        var objectsOcc = selectedModels[i].ObjOccListFilter()
        for (var j = 0; j < objectsOcc.length; j++) {

            if (objectsOcc[j].ObjDef().Attribute(Constants.AT_DOCUMENTATION, nLocale).getExternalDocument().isValid() == true) {
                objectsDefTemp.push(objectsOcc[j].ObjDef())
            }
            objectsOcc[j] = null;
        }
        objectsOcc = null;
    }
    var objectsDefTemp = ArisData.Unique(objectsDefTemp)
    return objectsDefTemp
}

function getADSLinks(selectedModels) {
    var tempLinks = new Array()

    for (var i = 0; i < selectedModels.length; i++) {

        if (selectedModels[i].Attribute(Constants.AT_ADS_LINK_1, nLocale).IsMaintained()) {
            var adsAtt = selectedModels[i].Attribute(Constants.AT_ADS_LINK_1, nLocale)
            tempLinks.push(adsAtt.getValue())
        }
        if (selectedModels[i].Attribute(Constants.AT_ADS_LINK_2, nLocale).IsMaintained()) {
            var adsAtt = selectedModels[i].Attribute(Constants.AT_ADS_LINK_2, nLocale)
            tempLinks.push(adsAtt.getValue())
        }
        if (selectedModels[i].Attribute(Constants.AT_ADS_LINK_3, nLocale).IsMaintained()) {
            var adsAtt = selectedModels[i].Attribute(Constants.AT_ADS_LINK_3, nLocale)
            tempLinks.push(adsAtt.getValue())
        }
        if (selectedModels[i].Attribute(Constants.AT_ADS_LINK_4, nLocale).IsMaintained()) {
            var adsAtt = selectedModels[i].Attribute(Constants.AT_ADS_LINK_4, nLocale)
            tempLinks.push(adsAtt.getValue())
        }

        var objectsOcc = selectedModels[i].ObjOccListFilter()

        for (var j = 0; j < objectsOcc.length; j++) {
            if (objectsOcc[j].ObjDef().Attribute(Constants.AT_ADS_LINK_1, nLocale).IsMaintained()) {
                var adsAtt = objectsOcc[j].ObjDef().Attribute(Constants.AT_ADS_LINK_1, nLocale)
                tempLinks.push(adsAtt.getValue())
            }
            if (objectsOcc[j].ObjDef().Attribute(Constants.AT_ADS_LINK_2, nLocale).IsMaintained()) {
                var adsAtt = objectsOcc[j].ObjDef().Attribute(Constants.AT_ADS_LINK_2, nLocale)
                tempLinks.push(adsAtt.getValue())
            }
            if (objectsOcc[j].ObjDef().Attribute(Constants.AT_ADS_LINK_3, nLocale).IsMaintained()) {
                var adsAtt = objectsOcc[j].ObjDef().Attribute(Constants.AT_ADS_LINK_3, nLocale)
                tempLinks.push(adsAtt.getValue())
            }
            if (objectsOcc[j].ObjDef().Attribute(Constants.AT_ADS_LINK_4, nLocale).IsMaintained()) {
                var adsAtt = objectsOcc[j].ObjDef().Attribute(Constants.AT_ADS_LINK_4, nLocale)
                tempLinks.push(adsAtt.getValue())
            }
            objectsOcc[j] = null;
        }
        objectsOcc = null;
    }
    var tempLinks = ArisData.Unique(tempLinks)
    return tempLinks
}

function jSonTextLabel(content, xPos, yPos) {
    var obj = new org.json.JSONObject();
    var contentArray = new org.json.JSONArray(content);
    obj.put("content", contentArray);
    obj.put("xPos", xPos);
    obj.put("yPos", yPos);

    return obj.toString(2);
}

function jSonGfxRoundedRectangleObjects(effect3D, brushColor, fillColor, height, width, xPos, yPos, zOrder, type, roundnessX, roundnessY, shaded) {

    var obj = new org.json.JSONObject();
    obj.put("effect3D", effect3D);
    obj.put("brushColor", brushColor);
    obj.put("fillColor", fillColor);
    obj.put("height", height);
    obj.put("width", width);
    obj.put("xPos", xPos);
    obj.put("yPos", yPos);
    obj.put("zOrder", zOrder);
    obj.put("type", type);
    obj.put("roundnessX", roundnessX);
    obj.put("roundnessY", roundnessY);
    obj.put("shaded", yPos);

    return obj.toString(2);

}

function jSonGfxPolygonObjects(effect3D, brushColor, fillColor, height, width, xPos, yPos, zOrder, type, pathArray, filled) {

    var obj = new org.json.JSONObject();

    var pathArray = new org.json.JSONArray(pathArray);

    obj.put("effect3D", effect3D);
    obj.put("brushColor", brushColor);
    obj.put("fillColor", fillColor);
    obj.put("height", height);
    obj.put("width", width);
    obj.put("xPos", xPos);
    obj.put("yPos", yPos);
    obj.put("zOrder", zOrder);
    obj.put("type", type);
    obj.put("path", pathArray);
    obj.put("filled", filled);

    return obj.toString(2);
}

function extractPictures(selectedModels) {
    for (var i = 0; i < selectedModels.length; i++) {
        var oleObjects = selectedModels[i].ComObjOccs()

        for (var j = 0; j < oleObjects.length; j++) {
            var imageGUID = getGUIDFromString(oleObjects[j].ObjectID())
            var image = oleObjects[j].Image();
            var format = image.getPreferredImageFormat()
            var imageByteData = image.getImageData(format)
            var imageName = imageGUID + "." + format
            var zipName = imageGUID + ".zip"

            zipDocFile(imageByteData, picPath, imageName, zipName)
        }
    }
}

function jSonPictures(imageGUID, imageHeight, imageWidth, imageX, imageY, imageZOrder) {

    var obj = new org.json.JSONObject();

    obj.put("imageGUID", imageGUID);
    obj.put("height", imageHeight);
    obj.put("width", imageWidth);
    obj.put("xPos", imageX);
    obj.put("yPos", imageY);
    obj.put("imageZOrder", imageZOrder);

    return obj.toString(2);
}

function getLanguageList() {
    var allLocales = new Array();
    var arrLanguageList = ArisData.getActiveDatabase().LanguageList();

    for (var i = 0; i < arrLanguageList.length; i++) {
        allLocales.push(arrLanguageList[i].LocaleId())
    }
    return allLocales
}

function getLocaleName(locale) {
    var arrLanguageList = ArisData.getActiveDatabase().LanguageList();
    for (var i = 0; i < arrLanguageList.length; i++) {
        if (locale == arrLanguageList[i].LocaleId()) {
            return arrLanguageList[i].Name(Context.getSelectedLanguage())
        }
    }
    return "language-independent"
}

function getContentLang(textOcc) {
    var contentArray = new org.json.JSONArray();
    var allLocales = getLanguageList()
    for (var i = 0; i < allLocales.length; i++) {
        var value = textOcc.TextDef().Attribute(Constants.AT_NAME, allLocales[i]).getValue()
        var language = allLocales[i]
        var languageName = getLocaleName(allLocales[i])
        if (value != "")
            contentArray = jSonObjCxnAttValuesLang(language, languageName, value, contentArray);
    }
    return contentArray.toString(2)
}