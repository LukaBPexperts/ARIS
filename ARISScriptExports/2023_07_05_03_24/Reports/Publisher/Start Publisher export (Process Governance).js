/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

/*
*   Script to create a Business Publisher export
*   Author: Software AG
*   Date: 08/01/2007
*/


var oBusinessPublisher=Context.getComponent("BusinessPublisher");

function main()
{
    
    
    
    var sBPResult;
    // Extract Business Publisher script component
    
    if (oBusinessPublisher==null) {
        return;
    }
    else {
        var sBusinessPublisherServerName;
        var sExportName;
        var sExportDescription;

        var sExportUser;
        var sExportPassword;
        
        // web application server adress (IP or full qualified name)
        if (Context.getProperty("bp_ServerLocation") != null) {
            sBusinessPublisherServerName = Context.getProperty("bp_ServerLocation");
        } else {
            sBusinessPublisherServerName = "localhost";
        }
        
        // export name
        if (Context.getProperty("bp_ExportName") != null) {
            sExportName = Context.getProperty("bp_ExportName").replace(" ", "");
        } else {
            sExportName = getString("BP_2");
        }

        // export description                
        if (Context.getProperty("bp_ExportDesc")) {
            sExportDescription = Context.getProperty("bp_ExportDesc");
        } else {
            sExportDescription = getString("BP_3");
        } 
        
        // export user
        if (Context.getProperty("bp_ExportUser") != null) {
            sExportUser = Context.getProperty("bp_ExportUser");
        } else {
            sExportUser = getString("BP_10");
        }
                
        // export password
        if (Context.getProperty("bp_ExportPassword") != null) {
            sExportPassword = Context.getProperty("bp_ExportPassword");
        } else {
            sExportPassword = getString("BP_10");
        }
                
        //Select the active database to create export component instances.
        var databaseToexport = ArisData.getActiveDatabase();
        if (Context.getProperty("BP_ACTIVATE") != null && Context.getProperty("BP_ACTIVATE") == "TRUE") {
            oBusinessPublisher.activateExport(sBusinessPublisherServerName, sExportName, true);
        }
        else if (Context.getProperty("BP_DEACTIVATE") != null && Context.getProperty("BP_DEACTIVATE") == "TRUE") {
            oBusinessPublisher.activateExport(sBusinessPublisherServerName, sExportName, false);
        }
        else if (Context.getProperty("BP_DELETE") != null && Context.getProperty("BP_DELETE") == "true") {
            oBusinessPublisher.deleteExport(sBusinessPublisherServerName, sExportName);
        }
        else if (Context.getProperty("BP_UPDATE") != null && Context.getProperty("BP_UPDATE") == "TRUE") {
            if (oBusinessPublisher.exportExists(sBusinessPublisherServerName, sExportName)) {
                exportData = oBusinessPublisher.getExportData(databaseToexport, sBusinessPublisherServerName, sExportName);
                if ((Context.getProperty("bp_Changelist") != null) && (databaseToexport.isVersionable())) {
                    try {
                        exportData.setASN(parseInt(Context.getProperty("bp_Changelist")));            
                    } catch (e) {
                        exportData.setASN(-1);    
                    }
                } else {
                    exportData.setASN(-1);
                }
                oBusinessPublisher.updateExport(databaseToexport, sBusinessPublisherServerName, sExportName, exportData.getASN());
                setExportToAnonymous(oBusinessPublisher);
            }
        }
        else /*(Context.getProperty("BP_CREATE") != null && Context.getProperty("BP_CREATE") == "TRUE")*/ {
            var exportData;
            if (oBusinessPublisher.exportExists(sBusinessPublisherServerName, sExportName)) {
                exportData = oBusinessPublisher.getExportData(databaseToexport, sBusinessPublisherServerName, sExportName);
                if ((Context.getProperty("bp_Changelist") != null) && (databaseToexport.isVersionable())) {
                    try {
                        exportData.setASN(parseInt(Context.getProperty("bp_Changelist")));            
                    } catch (e) {
                        exportData.setASN(-1);    
                    }
                } else {
                    exportData.setASN(-1);
                }
                oBusinessPublisher.updateExport(databaseToexport, sBusinessPublisherServerName, sExportName, exportData.getASN());
            }
            else {
                //Create export component instances.
                exportData = oBusinessPublisher.createABPExportDataInstance(databaseToexport);

                //set exportname and description
                exportData.setExportName(sExportName);
                exportData.setExportDescription(sExportDescription);
                
                //Check the amount of existing user groups in database
                //If no user group exists, it make no sense to create export
                if (databaseToexport.UserGroupList().length == 0) {
                    return;
                }
            
                exportData.setWebAppServerInfo(exportData.createIBPWebAppServerInfoInstance(sBusinessPublisherServerName));
        
                // Changelist
                if ((Context.getProperty("bp_Changelist") != null) && (databaseToexport.isVersionable())) {
                    try {
                        exportData.setASN(parseInt(Context.getProperty("bp_Changelist")));            
                    } catch (e) {
                        exportData.setASN(-1);    
                    }
                } else {
                    exportData.setASN(-1);
                }
        
                // Associated export
                if (Context.getProperty("bp_AssociatedExport") != null) {
                    exportData.setAssociatedExport(Context.getProperty("bp_AssociatedExport"));            
                } else {
                    exportData.setAssociatedExport("");
                }
                
                // Template
                if (Context.getProperty("bp_Templateguid") != null) {
                	var templ = Context.getProperty("bp_Templateguid");
                    exportData.setSelectedTemplateAsString(templ);            
                } else {
                    exportData.setSelectedTemplateAsString("");
                }                
                //Language Packages
                var selectedLanguagePackages = getLanguagePackages(exportData, oBusinessPublisher); 
                if (selectedLanguagePackages == null || selectedLanguagePackages.length == 0) {
                    return;
                }
                exportData.setSelectedLanguagePackages(selectedLanguagePackages);
        
                
                //Default language        
                if (Context.getProperty("bp_DefaultLanguage") != null) {
                    var languagePackage = exportData.createIBPLanguagePackageDataInstance();
                    var defaultLanguage = getLocaleIDFromLanguageAndCountry(Context.getProperty("bp_DefaultLanguage"));
                    exportData.setDefaultLanguageId(defaultLanguage);            
                } else {
                    //Set first language to default language
                    exportData.setDefaultLanguageId(selectedLanguagePackages[0].getDatabaseLocaleId());    
                }
            
                // No evaluation filter
                if ((Context.getProperty("bp_Evalfilterguid") != null) && (Context.getProperty("bp_Evalfilterguid").length > 0)) {
                	var evalFilter = Context.getProperty("bp_Evalfilterguid");
                    exportData.setSelectedEvalFilterAsString(evalFilter);
                } else {
                    exportData.setSelectedEvalFilterAsString("");
                }        
                //All object types allowed
                exportData.setAllowedObjectTypes(databaseToexport.ActiveFilter().ObjTypes());
        
                
                //In this case both export will be exported with the same layout
                if (Context.getProperty("bp_Layout") != null) {
                    exportData.setSelectedLayout(Context.getProperty("bp_Layout"));            
                } else {
                    var layoutinfos = oBusinessPublisher.getAvailableLayouts(sBusinessPublisherServerName);
                    exportData.setSelectedLayout(layoutinfos[0].getLayoutName());
                }
        
				//not used anymore
                //Method filter with format: GUID1;GUID2;.... 
                //if (Context.getProperty("bp_MethodFilters") != null) {       
                //   var selectedMethodFilters = new Array();
        
                //   var tokenizer = new java.util.StringTokenizer(Context.getProperty("bp_MethodFilters"), ";");
                //    while(tokenizer.hasMoreTokens()) {
                //        var filterGUID = aGUIDFactory.createGUID(tokenizer.nextToken());
                //        selectedMethodFilters.push(filterGUID);
                //    }
                //    exportData.setSelectedMethodFilters(selectedMethodFilters);
                //} else {
                //    exportData.setSelectedMethodFilters(new Array(aGUIDFactory.getNULLGUID()));            
                //}
        
                //Profile -  Both exports will get the same graphical properties
                var profileData = fillProfileData(exportData.getProfileData());     
                exportData.setProfileData(profileData);
        
                //Create export
                oBusinessPublisher.createExport(databaseToexport, exportData);
            }

            //Set export profile to anonymous Please specify user group guid and filter guid and if anonymous is true or false.
            setExportToAnonymous(oBusinessPublisher);    
                
            //Activate Export
            oBusinessPublisher.activateExport(sBusinessPublisherServerName, sExportName, true);
            
            var sBPResult = createBPURL(exportData, getBoolPropertyValue("bp_Anonym"));
            //Context.setProperty("bp_resultURL", sBPResult);

            var ooutfile = Context.createOutputObject (); // Output object is created
            ooutfile.OutputTxt("<result>\r\n");
            ooutfile.OutputTxt("\t<bpUrl>" + sBPResult + "</bpUrl>\r\n");
            ooutfile.OutputTxt("</result>\n\n");
            ooutfile.WriteReport(); //This line only tests the result
        }
    }
}


function setExportToAnonymous(p_oBusinessPublisher) {
    
    var bExportIsAnonym = getBoolPropertyValue("bp_Anonym");
    
    var sBusinessPublisherServerName = Context.getProperty("bp_ServerLocation");
    if (sBusinessPublisherServerName == null) {
        sBusinessPublisherServerName = "localhost";
    }
    
    var sExportName = Context.getProperty("bp_ExportName").replace(" ", "");
    if (sExportName == null) {
        sExportName = getString("BP_2");
    }   
    
    var sUserGroupName = Context.getProperty("bp_UserGroupName");
    //Only test
    if (sUserGroupName == null) {
        sUserGroupName = "arisservice";
    }
    
    var databaseToexport = ArisData.getActiveDatabase();
    var findresult = databaseToexport.Find( Constants.SEARCH_USERGROUP, 1, Constants.LCID_NEUTRAL, sUserGroupName, Constants.SEARCH_CMP_EQUAL )
    if (findresult == null || findresult.length == 0) {
        return;
    }
    var sUserGroupGuid = findresult[0].GUID();
    
    var sFilterGuid;
    if (Context.getProperty("bp_MethodFilters") != null) {       
        var selectedMethodFilters = new Array();

        var tokenizer = new java.util.StringTokenizer(Context.getProperty("bp_MethodFilters"), ";");
        while(tokenizer.hasMoreTokens()) {
            //choose the first guid
            sFilterGuid = tokenizer.nextToken();
            break;
        }
    } else {
        sFilterGuid = databaseToexport.ActiveFilter().FilterKey();
    }
    p_oBusinessPublisher.setProfileToAnonymous(sBusinessPublisherServerName, sExportName, sUserGroupGuid, sFilterGuid, bExportIsAnonym);   
}

function fillProfileData(profileData) {
    try {
        if (Context.getProperty("bp_Profilename") != null) {
            profileData.setProfileName(Context.getProperty("bp_Profilename"));
        } else {
            profileData.setProfileName(getString("BP_8"));    
        }
            
        if (Context.getProperty("bp_Profiledesc") != null) {
            profileData.setProfileDescription(Context.getProperty("bp_Profiledesc"));
        } else {
            profileData.setProfileDescription(getString("BP_9"));
        }
        
        if (Context.getProperty("bp_UsePrintScale") != null) {
            if (Context.getProperty("bp_UsePrintScale") == "TRUE") {
                profileData.setUsePrintScale(true);
            } else {
                profileData.setUsePrintScale(false);
            }
        } else {
            profileData.setUsePrintScale(false);
        }
        
        if (Context.getProperty("bp_InitialScale") != null) {
            profileData.setInitialScale(Context.getProperty("bp_InitialScale"));    
        } else {
            profileData.setInitialScale(100);
        }

        if (Context.getProperty("bp_ZoomSteps") != null) {
            var zoomsteps = new Array();
            var tokenizer = new java.util.StringTokenizer(Context.getProperty("bp_ZoomSteps"), ",");
            while(tokenizer.hasMoreTokens()) {
                var step = tokenizer.nextToken();
                zoomsteps.push(parseInt(step));
            }
            profileData.setScaleValues(zoomsteps);
        } else {
            profileData.setScaleValues(new Array(25, 50, 75, 125, 150));   
        }

        if (Context.getProperty("bp_BlackWhite") != null) {
            if (Context.getProperty("bp_BlackWhite") == "TRUE") {
                profileData.setBlackWhite(true);                    
            } else {
                profileData.setBlackWhite(false);                    
            }
        } else {
            profileData.setBlackWhite(false);    
        }


        if (Context.getProperty("bp_Transparent") != null) {
            if (Context.getProperty("bp_Transparent") == "TRUE") {
                profileData.setTransparent(true);                    
            } else {
                profileData.setTransparent(false);                    
            }
        } else {
            profileData.setTransparent(false);    
        }

        //Sorting alphabetical
        if (Context.getProperty("bp_ModelSort") != null) {
            profileData.setModelSort(Context.getProperty("bp_ModelSort"));
        } else {
            profileData.setModelSort(0);
        }
        if (Context.getProperty("bp_ObjectSort") != null) {
            profileData.setObjectSort(Context.getProperty("bp_ObjectSort"));
        } else {
            profileData.setObjectSort(0);
        }

        //Sorting methodical
        if (Context.getProperty("bp_AttributeSort") != null) {
            profileData.setAttributeSort(Context.getProperty("bp_AttributeSort"));
        } else {
            profileData.setAttributeSort(2);
        }

        //List of attribute type nums        
        if (Context.getProperty("bp_LinkList") != null) {
            var linklist = new Array();
            var tokenizer = new java.util.StringTokenizer(Context.getProperty("bp_LinkList"), ";");
            while(tokenizer.hasMoreTokens()) {
                var attrTypeNum = tokenizer.nextToken();
                linklist.push(parseInt(attrTypeNum));
            }
            profileData.setLinkList(linklist);
        } else {
            profileData.setLinkList(new Array(Constants.AT_LINK, Constants.AT_EXT_1, Constants.AT_EXT_2, Constants.AT_EXT_3));
        }

        if (Context.getProperty("bp_CopyDocs") != null) {
            if (Context.getProperty("bp_CopyDocs") == "TRUE") {
                profileData.setCopyDocuments(true);
            } else {
                profileData.setCopyDocuments(false);
            }
        } else {
            profileData.setCopyDocuments(false);
        }
        
        if (Context.getProperty("bp_TempFN") != null) {
            if (Context.getProperty("bp_TempFN") == "TRUE") {
                profileData.setTempFNames(true);
            } else {
                profileData.setTempFNames(false);
            }
        } else {
            profileData.setTempFNames(false);
        }
    } catch (e) {}
    return profileData;
}

function getLanguagePackages(exportData, oBusinessPublisher) {

    //Create language combinations
    var selectedLanguagePackages = new Array();
    if (Context.getProperty("bp_ExportLanguages") != null) {
        var packageTokenizer = new java.util.StringTokenizer(Context.getProperty("bp_ExportLanguages"), ";");
        if (!packageTokenizer.hasMoreTokens()) {
            return null;
        }
        var langIdx = 0;
        while(packageTokenizer.hasMoreTokens()) {
            var languagePackage = exportData.createIBPLanguagePackageDataInstance();
            languagePackage.setPackageName(getString("BP_6") + langIdx);
            languagePackage.setPackageDescription(getString("BP_7") + langIdx++);

            var language = packageTokenizer.nextToken();
            var localeId = getLocaleIDFromLanguageAndCountry(language);
            languagePackage.setInterfaceLocaleId(localeId);
            languagePackage.setDatabaseLocaleId(localeId);
            languagePackage.setMethodLocaleId(localeId);

            selectedLanguagePackages.push(languagePackage);
        }
    } else {
        var availableInterfaceLanguages = oBusinessPublisher.getAvailableInterfaceLanguages(exportData.getWebAppServerInfo().getServerAddress());
        if (availableInterfaceLanguages.length <= 0) {
            return null;
        }
        //Language combinations where interface, database- and methodlanguage are equal
        for ( var i = 0 ; i < availableInterfaceLanguages.length ; i++ ){
            var languagePackage = exportData.createIBPLanguagePackageDataInstance();
            languagePackage.setPackageName(getString("BP_6") + i);
            languagePackage.setPackageDescription(getString("BP_7") + i);
            languagePackage.setInterfaceLocaleId(availableInterfaceLanguages[i].getLocaleId());
            languagePackage.setDatabaseLocaleId(availableInterfaceLanguages[i].getLocaleId());
            languagePackage.setMethodLocaleId(availableInterfaceLanguages[i].getLocaleId());
            selectedLanguagePackages.push(languagePackage);
        }
    }
    return selectedLanguagePackages;
}

function createBPURL(exportData, p_bExportIsAnonym) {

    //constants
    var PARAM_LOGIN = "login=";
    var PARAM_PASSWORD = "password=";
    var PARAM_LOCALEID = "localeid=";
    var PARAM_FILTERGUID = "filterguid=";
    var PARAM_EXPORTNAME = "exportname=";
    var PARAM_USERGROUPGUID = "usergroupguid=";
    var HTTP = "http://";
    var CONTEXTWITHACTION = "/businesspublisher/link.do"
    var PARAM1 = "?";
    var PARAM2 = "&";
    
    var sBPUrl; 
            
    var sServerAddress = Context.getProperty("bp_ServerLocation");    
    var sPort = Context.getProperty("bp_ServerPort");    
    if (sPort == null || sPort == 0) {
        sPort = "19990";
    }
    var sExportName = Context.getProperty("bp_ExportName").replace(" ", "");
    if (sExportName == null) {
        sExportName = getString("BP_2");
    }   
    
    var sUserGroupName = Context.getProperty("bp_UserGroupName");
    if (sUserGroupName == null) {
        return;   
    }
    
    var languagePackage = exportData.createIBPLanguagePackageDataInstance();
    var databaseToexport = ArisData.getActiveDatabase();
    var findresult = databaseToexport.Find( Constants.SEARCH_USERGROUP, 1, Constants.LCID_NEUTRAL, sUserGroupName, Constants.SEARCH_CMP_EQUAL )
    var sUserGroupGuid = findresult[0].GUID();
    var sLocaleId = getLocaleIDFromLanguageAndCountry(Context.getProperty("bp_ExportLocale"));
    var sUser = Context.getProperty("bp_UserName") != null ? Context.getProperty("bp_UserName") : "";
    var sPassword = Context.getProperty("bp_Password") != null ? Context.getProperty("bp_Password") : "";
    
    if (p_bExportIsAnonym) {
        sUser = "anonymous";
        sPassword = "anonymous";    
    }
    
    var sFilterGuid;
    if (Context.getProperty("bp_MethodFilters") != null) {       
        var selectedMethodFilters = new Array();

        var tokenizer = new java.util.StringTokenizer(Context.getProperty("bp_MethodFilters"), ";");
        while(tokenizer.hasMoreTokens()) {
            //choose the first guid
            sFilterGuid = tokenizer.nextToken();
            break;
        }
    } else {
        sFilterGuid = databaseToexport.ActiveFilter().FilterKey();
    }
    
    
    //http://tlabapp9srv:9090/businesspublisher/link.do?login=anonymous&password=anonymous&filterguid*******&usergroupguid=******&exportname=*****
    
    sBPUrl = HTTP + sServerAddress + ":" + sPort + CONTEXTWITHACTION + PARAM1 + 
        PARAM_LOGIN + sUser + PARAM2 +
        PARAM_PASSWORD + sPassword + PARAM2 +
        PARAM_EXPORTNAME + sExportName + PARAM2 + 
        PARAM_LOCALEID + sLocaleId + PARAM2 + 
        PARAM_USERGROUPGUID + sUserGroupGuid + PARAM2 +
        PARAM_FILTERGUID + sFilterGuid;
        
    return getBoolPropertyValue("bp_needURLEnc") ? oBusinessPublisher.removeSpecialCharacters(sBPUrl) : sBPUrl;
}

function getBoolPropertyValue(p_sPropKey) {
    var property = Context.getProperty(p_sPropKey);
    if (property != null) {
        return (StrComp(property, "true") == 0);
    }
    return false;
}

function getLocaleIDFromLanguageAndCountry(p_languageAndCountry) {    
    var languageAndCountry = p_languageAndCountry.trim();
    if (languageAndCountry.indexOf("_") > 0) {
        return oBusinessPublisher.getLocaleId(org.apache.commons.lang.LocaleUtils.toLocale(languageAndCountry));
    } else {
        return oBusinessPublisher.getLocaleId(languageAndCountry);
    }
}


main();
