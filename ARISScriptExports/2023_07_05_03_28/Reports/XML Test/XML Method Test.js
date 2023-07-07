
//This will define the Array.from method if it is not already present, allowing you to use it in your code
//Alternatively, you can use a JavaScript library that includes a polyfill for Array.from, such as the polyfill.io service. 
if (!Array.from) {
  Array.from = function(arrayLike) {
    return [].slice.call(arrayLike);
  };
}


//const xmlString = String.fromCharCode.apply(null, xmlBytes);

var nLocale = Context.getSelectedLanguage()

var selectedModels = ArisData.getSelectedModels()

var xmlAtt = selectedModels[0].Attribute(Constants.AT_DESC, nLocale).getValue()


var byteArrray = xmlAtt.getBytes()
var xmlReader = Context.getXMLParser(byteArrray)




/* var initXml = "<Assessments></Assessments>"
var bytes = (new java.lang.String(initXml)).getBytes();	
var xmlReader = Context.getXMLParser(bytes) */



var rootElement = xmlReader.getRootElement()		
var children = rootElement.getChildren()	

for (var i = 0; i < children.size(); i++) {
    var element = children.get(i);


    var a=2
  
}





xmlReader.getRootElement()





var xmlWriter = Context.createXMLOutputObject("settings.xml", "abc")

//const xmlDoc = Context.parseFromString("<test>", "text/xml");

var xmlReader = Context.getXMLParser(xmlData)




var a=2



addNewAssessment("xmlString", "assessmentName", "dateStarted", "dateCompleted", "strOutcome", "bMandatory")









function addNewAssessment(xmlString, assessmentName, dateStarted, dateCompleted, strOutcome, bMandatory) {
  // Parse the XML string into a DOM object
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(xmlString, "text/xml");

  // Check if the root Assessments node exists, create it if it doesn't
  var assessmentsNode = xmlDoc.getElementsByTagName("Assessments")[0];
  if (!assessmentsNode) {
    assessmentsNode = xmlDoc.createElement("Assessments");
    xmlDoc.appendChild(assessmentsNode);
  }

  // Check if the assessmentName node exists, create it if it doesn't
  var assessmentNode = xmlDoc.getElementsByTagName(assessmentName)[0];
  if (!assessmentNode) {
    assessmentNode = xmlDoc.createElement(assessmentName);
    assessmentsNode.appendChild(assessmentNode);
  }

  // Update the assessmentNode with the provided data
  assessmentNode.setAttribute("name", assessmentName);

  var attributeNode = xmlDoc.createElement("Attribute");
  attributeNode.setAttribute("name", "DateStarted");
  attributeNode.setAttribute("Value", dateStarted);
  assessmentNode.appendChild(attributeNode);

  attributeNode = xmlDoc.createElement("Attribute");
  attributeNode.setAttribute("name", "DateCompleted");
  attributeNode.setAttribute("Value", dateCompleted);
  assessmentNode.appendChild(attributeNode);

  attributeNode = xmlDoc.createElement("Attribute");
  attributeNode.setAttribute("name", "Outcome");
  attributeNode.setAttribute("Value", strOutcome);
  assessmentNode.appendChild(attributeNode);

  attributeNode = xmlDoc.createElement("Attribute");
  attributeNode.setAttribute("name", "Mandatory");
  attributeNode.setAttribute("Value", bMandatory);
  assessmentNode.appendChild(attributeNode);

  // Return the updated XML string
  return (new XMLSerializer()).serializeToString(xmlDoc);
}
