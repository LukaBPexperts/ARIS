/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

var REPORT_SEMANTICSCHECK_HIERARCHYSTRUCTURE        = "844c6cb0-c480-11da-14e3-00123f5c9b19";
var REPORT_SEMANTICSCHECK_RISK                      = "d9cd78b0-c63c-11da-14e3-00123f5c9b19";
var REPORT_SEMANTICSCHECK_CONTROL                   = "31f6bee0-c618-11da-14e3-00123f5c9b19";
var REPORT_SEMANTICSCHECK_TESTDEFINITION            = "a210dab0-c575-11da-14e3-00123f5c9b19";
var REPORT_SEMANTICSCHECK_USERGROUP                 = "d2e3ee40-c8c0-11da-14e3-00123f5c9b19";
var REPORT_SEMANTICSCHECK_USER                      = "c6f82c40-c8c0-11da-14e3-00123f5c9b19";
var REPORT_SEMANTICSCHECK_QUESTIONNAIRE_TEMPLATE    = "27cca360-49e3-11de-5f40-00237de75f8a";
var REPORT_SEMANTICSCHECK_AUDIT_TEMPLATE            = "3571b490-adf2-11e0-09f3-005056c00008";
var REPORT_SEMANTICSCHECK_POLICY_DEFINITION         = "8c0e95e0-6d21-11e1-1b6c-782bcb1f48e1";
 
var OUTPUTFILENAME = Context.getSelectedFile();

var HEADER = "***************************************************************************************\r\n"+
             "***************************************************************************************\r\n\r\n"+
             "      " + getString("TEXT_2")+
             "\r\n\r\n***************************************************************************************\r\n"+
             "***************************************************************************************\r\n\r\n\r\n";

             
function main() {
    var selection = ArisData.getSelectedGroups();
    if (selection.length == 0) {
        var databases = ArisData.getSelectedDatabases();
        if (databases.length != 0) {
            selection = new Array();
            selection.push(databases[0].RootGroup());
        }
    }
    
    //no selection -> nothing to do
    if (selection.length == 0) {
        return;
    }
    //approach not determined -> nothing to do
    if (determineApproach() == false) {
        return;    
    }

    determineModules();
    
    var reportComponent = Context.getComponent("Report");
    var aPropertyNames = null;
    var szOutput = "";
    
    var reports = new Array( REPORT_SEMANTICSCHECK_HIERARCHYSTRUCTURE
                             ,REPORT_SEMANTICSCHECK_RISK
                             ,REPORT_SEMANTICSCHECK_CONTROL
                             ,REPORT_SEMANTICSCHECK_TESTDEFINITION
                             ,REPORT_SEMANTICSCHECK_USERGROUP
                             ,REPORT_SEMANTICSCHECK_USER
                             ,REPORT_SEMANTICSCHECK_QUESTIONNAIRE_TEMPLATE
							 ,REPORT_SEMANTICSCHECK_AUDIT_TEMPLATE
                             ,REPORT_SEMANTICSCHECK_POLICY_DEFINITION
                             );
    var displayInfo = new Array( "DISPLAY_SC_HIERARCHYSTRUCTURE"
                                 ,"DISPLAY_SC_RISK"
                                 ,"DISPLAY_SC_CONTROL"
                                 ,"DISPLAY_SC_TESTDEFINITION"
                                 ,"DISPLAY_SC_USERGROUP"
                                 ,"DISPLAY_SC_USER"
                                 ,"DISPLAY_SC_QUESTIONNAIRE_TEMPLATE"
                                 ,"DISPLAY_SC_AUDIT_TEMPLATE"
                                 ,"DISPLAY_SC_POLICY_DEFINITION"
                                 );
    
    //execute the sub reports
    for (var i=0; i<reports.length; i++) {
        var reportExecInfo = reportComponent.createExecInfo(reports[i], selection, Context.getSelectedLanguage(), Constants.OutputTXT, Context.getSelectedFile());
        
        //set the approach info to the report context
        reportExecInfo.setProperty( "ARCM_APPROACH", g_approach );   
   
        //set information that report is executed by master report (formerly macro, hence the old flag name)
        reportExecInfo.setProperty( "aam_semcheck", "evalByMacro" );
        
        //pass classification properties from previous reports to the current report
        if (aPropertyNames != null && result != null) {
            for (var j=0; j<aPropertyNames.length; j++) {
                reportExecInfo.setProperty( aPropertyNames[j], result.getProperty(aPropertyNames[j]) );
            }
        }
        
        //pass output from previous reports to the current report
        if (i == (reports.length-1) ){
            reportExecInfo.setProperty( "reportdata", szOutput );   
        }
        else{
            reportExecInfo.setProperty( "reportdata", "" );
        }
        
        //write some execution info for the displayed dialog
        var indexInfo = new java.lang.String(getString("TEXT_3")).replaceFirst("%0", (i+1)).replaceFirst("%1", (displayInfo.length + 1));
        Context.writeStatus(getString(displayInfo[i]) + " (" + indexInfo + ")");
        
        //execute                
        result = reportComponent.execute(reportExecInfo);

        // Error handling
        
        // to use function from arcm-common.js
        if(result.getProperty("exception")) {
            var message = result.getProperty("scriptName") + ":\n" + result.getProperty("exception");
            Dialogs.MsgBox(message, Constants.MSGBOX_ICON_ERROR, getString("TEXT_ERROR_MESSAGE_TITEL.DBT"));
            return;
        }
        
        //read resulting reportdata
        if ((result.getProperty("reportdata") != null) && (!result.getProperty("reportdata").equals(""))){            
            szOutput += result.getProperty("reportdata") + "\r\n\r\n\r\n";
        }
        
        //read classification properties
        aPropertyNames = result.getPropertyNames();
    }
    
    //after executing all reports - output file
    if (szOutput != null) {
        szOutput = HEADER + szOutput;
        writeErrorreport(szOutput, OUTPUTFILENAME);
    }
}             

main();