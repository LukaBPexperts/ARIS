/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

/**
 * Object performing the calculation for the QMS use case "Validity date"
 */
var ValidityDateUseCase = function() {
    const MILLIES_OF_ONE_DAY = 24 * 60 * 60 * 1000;
    const VALID_UNTIL_PROPERTY = "ValidityDate_ValidUntilProperty";
    const TIME_PERIOD = "ValidityDate_TimePeriod";
    
    const m_valueAccess = Context.getParameterValueAccess();
    const m_validUntilProperty = m_valueAccess.getParameterValue(VALID_UNTIL_PROPERTY);
    const m_timePeriodList = getTimePeriodList();
    
    
    const m_outputParameters = new OutputParameters("ValidityDate");
    const m_output = createXmlOutput(m_outputParameters);
    m_output.setColumns([[getString("TABLE_COLUMN_HEADER_NAME"), "text"], 
                         [getString("TABLE_COLUMN_HEADER_GUID"),"text"],
                         [getString("TABLE_COLUMN_HEADER_VALIDITY_DATE"), "text"],
                         [getString("TABLE_COLUMN_HEADER_PERIOD"), "text"]]);
    
    /**
     * Returns properties which are required by this use case
     * sQuery - Name of the query ("chain" or "process") (JavaScript String)
     * return - Array of Strings (JavaScript String) with the property names
     */
    this.getPropertyTree = function(sQuery) {
        
        return ["name", m_validUntilProperty];
    }
    
    /**
     * Performs calculation for one data node, called when the data tree is travesed.
     * For details see TreeGraph.traverse();
     * Object function(TreeGraphNode currentNode, java.util.Map<TreeGraphNode, Object> childNodeResults)
     * The provided node has as item a result object of the data query
     */
    this.processNodeList = function(nodeList)
    {
        for(var itNodes = nodeList.iterator(); itNodes.hasNext();)
        {
            var node = itNodes.next();
            processNode(node);
        }
    }
     
    processNode = function(node, childResults) {
        const type = node.getItemType();
        
        const resultEntry = node.getItem();
        
        if(resultEntry == null) {
            return;
        }

        const name = getName(resultEntry);
        const guid = getGuid(resultEntry);
        const validityDate = resultEntry.getProperty(m_validUntilProperty).getDate();
        const validityString = getValidityString(validityDate);
        const periodString = getPeriodString(validityDate);

        m_output.addRow([name, guid, validityString, periodString]);
    }

    function getTimePeriodList() {
        const list = new java.util.ArrayList(m_valueAccess.getParameterValue(TIME_PERIOD).getValueList());
        list.sort(null);
        return list;
    }
    
    function getName(resultEntry) {
        var name = resultEntry.getName();

        if(name == null) {
            name = new java.lang.String(getString("UNTITLED"));
        }
        
        name = name.replaceAll("\r\n", " ");
        name = name.replaceAll("\r", " ");
        name = name.replaceAll("\n", " ");
        
        return name;
    }
    
    function getGuid(resultEntry) {
        var guid = resultEntry.getGuid();

        if (guid == null) {
            guid = new java.lang.String(getString("NO_GUID"));
        }
        
        return guid;
    }
    
    function getValidityString(validityDate) {
        
        if (validityDate == null) {
            return getString("NOT_SPECIFIED");
        }
        
        const dateFormat = new java.text.SimpleDateFormat("yyyy-MM-dd");
        return dateFormat.format(validityDate);
    }
    
    function getPeriodString(validityDate) {
        
        if (validityDate == null) {
            return getString("UNDEFINED");
        }

        const currentDateInMillies = Date.now()
        const validityDateInMillies = validityDate.getTime() + MILLIES_OF_ONE_DAY - 1;
        const timeLeft = validityDateInMillies - currentDateInMillies;

        if (timeLeft < 0) {
            return getString("PERIOD_EXPIRED");
        }
        
        for (var it = m_timePeriodList.iterator(); it.hasNext();) {
            var timePeriod = it.next();
            
            if (timeLeft <= timePeriod * MILLIES_OF_ONE_DAY) {
                return formatstring1(getString("PERIOD_DAYS"), timePeriod);
            }
        }
        
        return getString("PERIOD_LATER");
    }

    /**
     * Called when the calculations are finished
     */    
    this.finish = function() {
        uploadXmlOutputToADS(m_output, m_outputParameters);
    }
}