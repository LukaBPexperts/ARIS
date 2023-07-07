/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

//
// provides 2 output objects:
// XMLOutputObject
// JSONOutputObject
//
// methods: 
// - constructor (no parameters)
// - setColumns(column array)
// - addRow(row items array) //items of 1 row, must have same size as |column array|
// - writeReport()
/** example
//var sampleOutput = new XMLOutputObject()
var sampleOutput = new JSONOutputObject()
sampleOutput.setColumns([["id","number"], ["name", "text"]]) //array: [name, type] valid values for type: "number", "text", "date"
sampleOutput.addRow(["1234567890", "my first name"])
sampleOutput.addRow(["0987654321", "my second name"])
sampleOutput.writeReport()
*/

// 2016-05-31 Additional parameter 'p_sFileName' added

var XMLOutputObject = function(p_sFileName, p_bUseSimpleFormat) {

    //private property
    var m_bUseSimpleFormat = p_bUseSimpleFormat
    var m_sFileName = p_sFileName
    
    Context.setSelectedFormat(Constants.OutputXML)
    var m_output = Context.createXMLOutputObject(p_sFileName, "amzresult")
    m_output.setEncoding("UTF-8")

    //private property
    var m_eRoot = m_output.getRootElement() 
    
    //private property
    var m_eHeader = new org.jdom.Element("metainfos");

    if(!m_bUseSimpleFormat)
        m_eRoot.addContent(m_eHeader);
    
    //private property
    var m_eData = new org.jdom.Element("data");
    //eRows.setAttribute(new Attribute("id", "1"));
    m_eRoot.addContent(m_eData);   
    
    //private property
    var m_aColumnNames;
    var m_aDatatypes;

    this.setColumns = function(aColumns)
    {
        m_aColumnNames = new Array()
        m_aDatatypes   = new Array()
        
        var eColumns = new org.jdom.Element("columns");
        m_eHeader.addContent(eColumns);

        for(var i=0; i<aColumns.length; i++)
        {
            m_aColumnNames.push(aColumns[i][0])
            m_aDatatypes.push(aColumns[i][1])
            
            var eColumn = new org.jdom.Element("column");
            eColumn.setAttribute(new org.jdom.Attribute("idx", ""+i));
            eColumn.setAttribute(new org.jdom.Attribute("datatype", aColumns[i][1]));
            eColumn.setAttribute(new org.jdom.Attribute("name", aColumns[i][0]));
            eColumn.setText( aColumns[i][0] )
            eColumns.addContent(eColumn);
        }
        
        if(!m_bUseSimpleFormat)
        {
            var eParameters = new org.jdom.Element("parameters");
            var eParameter = new org.jdom.Element("parameter");
            eParameter.setAttribute(new org.jdom.Attribute("datatype", "TEXT"));
            eParameter.setText( "param" )
            eParameters.addContent(eParameter);
            m_eHeader.addContent(eParameters);        
        }
    }
    
    this.addRow = function(aCellValues)
    {
        if(aCellValues.length != m_aColumnNames.length)
            throw "addRow: Incorrect column count " + aCellValues.length + ". Should be " + m_aColumnNames.length
        
        var eRow = new org.jdom.Element("row");
        for(var i=0; i<aCellValues.length; i++)
        {
            var eText;
            if(m_bUseSimpleFormat)
            {
                var sElementName = m_aColumnNames[i]
                while(sElementName.indexOf(" ")>=0)
                    sElementName = sElementName.replace(" ", "-")    
                eText = new org.jdom.Element(sElementName);
            }
            else
                eText = new org.jdom.Element(m_aDatatypes[i]);
            
            eText.setAttribute(new org.jdom.Attribute("idx", ""+i));
            eText.setText( aCellValues[i] )
            eRow.addContent(eText);
        }        
        m_eData.addContent(eRow);
    }
    
    this.writeReport = function() {
        m_output.WriteReport()
    }    
    
    this.getFileName = function() {
        return m_sFileName;
    }    

    this.getRootElement = function() {
        return m_eRoot;
    }
    
    this.cloneHeader = function() {
        return m_eHeader.clone();
    }
    
    this.getContentAsByteArray = function() {
        return m_output.getDocumentContent();
    }
}

var JSONOutputObject = function() {
    
    Context.setSelectedFile("result.json")
    Context.setSelectedFormat(Constants.OutputTXT)
    //private property
    var m_output = Context.createOutputObject()

    //private property
    var m_columns
    this.setColumns = function(aColumns)
    {
        m_columns = new Array()
        for(var i=0; i<aColumns.length; i++)
        {
            m_columns.push(aColumns[i][0])
        }
    }
    this.getColumns = function()
    {
        return m_columns;
    }
    
    //private property
    var m_rows = new Array()
    this.addRow = function(aCellValues)
    {
        m_rows.push(aCellValues)
    }
    
    this.writeReport = function() {
        var result = {"columns" : m_columns, "rows" : m_rows}
        m_output.OutputTxt ( JSON.stringify(result) ) 
        m_output.WriteReport()
    }
}
