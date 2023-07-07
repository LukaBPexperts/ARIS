var nLocale = Context.getSelectedLanguage()
var selectedModels = ArisData.getSelectedModels()

addNewAssessment(selectedModels[0], "nameAtt", "MONITORING", "testDate", "testCompleted", "testOutcome", true)


function addNewAssessment(objDef, attrName, assessmentName, dateStarted, dateCompleted, strOutcome, bMandatory) {
    var nLocale = Context.getSelectedLanguage()
    //var attNum = getAttributeNumber(objDef, attrName)

    //var xmlAtt = objDef.Attribute(attNum, nLocale) //commented for test

    //var xmlAttValue = xmlAtt.getValue() //commented for test

    var xmlAttValue = objDef.Attribute(Constants.AT_DESC, nLocale).getValue()

    var byteArrray = xmlAttValue.getBytes()
    var xmlReader = Context.getXMLParser(byteArrray)


    var rootElement = xmlReader.getRootElement()

    //if rootElement don't exist or name of the element not Assessments
    if (rootElement == null || rootElement.getName() != "Assessments") {
        var initXml = "<Assessments></Assessments>"
        var bytes = (new java.lang.String(initXml)).getBytes();

        var xmlReader = Context.getXMLParser(bytes)
        var rootElement = xmlReader.getRootElement()

    }

    var assessmentElement = rootElement.getChild(assessmentName)

    var xmlWriter = Context.createXMLOutputObject("temp.xml", xmlReader.getDocument())

    //check if assessmentName element exist, if not create it
    if (assessmentElement == null) {
        //create new assesment element
        var assessmentElement = xmlWriter.addElement(rootElement, assessmentName);

        //add attributes to this element
        var attributeElement = xmlWriter.addElement(assessmentElement, "DateStarted");
        attributeElement.setAttribute("Value", dateStarted);

        var attributeElement = xmlWriter.addElement(assessmentElement, "DateCompleted");
        attributeElement.setAttribute("Value", dateCompleted);

        var attributeElement = xmlWriter.addElement(assessmentElement, "Outcome");
        attributeElement.setAttribute("Value", strOutcome);

        var attributeElement = xmlWriter.addElement(assessmentElement, "Mandatory");
        attributeElement.setAttribute("Value", bMandatory);

    }
    else {
        //update attributes
        assessmentElement.removeChild("DateStarted")
        var attributeElement = xmlWriter.addElement(assessmentElement, "DateStarted");
        attributeElement.setAttribute("Value", dateStarted);

        assessmentElement.removeChild("DateCompleted")
        var attributeElement = xmlWriter.addElement(assessmentElement, "DateCompleted");
        attributeElement.setAttribute("Value", dateCompleted);

        assessmentElement.removeChild("Outcome")
        var attributeElement = xmlWriter.addElement(assessmentElement, "Outcome");
        attributeElement.setAttribute("Value", strOutcome);

        assessmentElement.removeChild("Mandatory")
        var attributeElement = xmlWriter.addElement(assessmentElement, "Mandatory");
        attributeElement.setAttribute("Value", bMandatory);
    }


    //WRITE XML BACK TO ATTRIBUTE
    var byteArr = xmlWriter.getDocumentContent()

    // Polyfill for Array.from
    if (!Array.from) {
        Array.from = function(arrayLike) {
            return [].slice.call(arrayLike);
        };
    }

    // Convert the array of bytes to a string
    var xmlString = String.fromCharCode.apply(null, Array.from(byteArr));

    var setAtt = objDef.Attribute(Constants.AT_DESC, nLocale).setValue(xmlString)

}