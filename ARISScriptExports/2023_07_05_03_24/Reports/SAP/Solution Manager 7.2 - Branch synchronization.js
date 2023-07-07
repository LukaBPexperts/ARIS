/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */


main();

function main()
{
    ROLE_CONSIDER_ORGUNITS = 1;
    ROLE_CONSIDER_POSITIONS = 2;
    ROLE_CONSIDER_JOBS = 4;
    ROLE_CONSIDER_USERS = 8;

    if (ArisData.getSelectedObjDefs().length == 1)
    {
        var SM72 = Context.getComponent("SM72");
        var b2ARIS = false;
        var options = SM72.createOptions(ArisData.getSelectedObjDefs()[0], b2ARIS);

        options.setSAPLogonUser("User");
        options.setSAPLogonPassword("Password");
        options.setSAPLogonLanguage("E");
        
        setDefaultSynchParams(options);
        
        var result = SM72.synchronize(ArisData.getSelectedObjDefs()[0], options);
        if (result.getLogFileName() != null && result.getLogFileName().length() > 0)
        {
            saveLogFile(result.getLogFileName());       
        }
        else
        {
            var ooutfile = Context.createOutputObject ();
            ooutfile.OutputTxt("<result>\r\n");
            ooutfile.OutputTxt("\t<success>" + result.getResult() + "</success>\r\n");
            ooutfile.OutputTxt("\t<error code=\"0\" message=\"" + result.getErrorMessage() + "\"/>\r\n");
            ooutfile.OutputTxt("</result>\n\n");
            ooutfile.WriteReport(); 
        }
    }  
}

function saveLogFile(sFnm) 
{    
    try 
    {   
        var inputStream = null;
        var source = new java.io.File(sFnm);
        inputStream = new java.io.FileInputStream(source);
        var output = Context.createOutputObject();
        Context.setSelectedFile(source.getName());
                       
        var buffer = java.lang.reflect.Array.newInstance(java.lang.Byte.TYPE, source.length());
        var i = inputStream.read(buffer);
        inputStream.close();
        inputStream = null;
        Context.setFile(source.getName(), Constants.LOCATION_OUTPUT, buffer)
    }
    catch(e) 
    {
        if(inputStream!=null) inputStream.close(); 
        return null;
    }
    output.WriteReport()
}

// default parameters for synchronization ARIS to SOLMAN
function setDefaultSynchParams(options)
{
    options.setTemplateGuid("2ed57d10-68c8-11d7-5d85-000bcd25c95f");
    options.setRelayoutFad(false);
    options.setRecursive(true);
    
    //structure
    options.setAddStructure(true);
    options.setDeleteStructure(true);
    options.setRenameStructure(true);
    options.setDeleteShortcutAndSlaveWithMaster(false);

    //tcodes
    options.setAddTCode(false);
    options.setDeleteTCode(false);
    
    //downloadable assignments
    options.setTransferMasterData(false);
    options.setTransferDocumentation(true);
    options.setTransferOrgUnits(false);
    options.setDeleteMasterData(false);
    options.setDeleteDocumentation(true);
    options.setDeleteOrgUnits(false);
    
   
    //weblinks
    options.setDeleteWebLink(false);
    options.setTransferWebLink(false);
    options.setWebLinkPrefix("");
}
