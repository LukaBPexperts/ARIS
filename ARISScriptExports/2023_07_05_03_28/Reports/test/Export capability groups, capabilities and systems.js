/* -------------
Normal variables for initating a script
------------- */
//Global variables
var g_nLoc = Context.getSelectedLanguage();

var g_entireMethodGUID = "dd838074-ac29-11d4-85b8-00005a4053ff";

//Getting user-defined attributes
var g_method = ArisData.ActiveFilter();

//Attribute capability level
var g_attrCapabilityLvl = g_method.UserDefinedAttributeTypeNum("149d8a20-2def-11e4-7f33-b4b52f5aab7c");


/* -------------
Global variables for writing to Excel
------------- */
var g_exc = Context.createExcelWorkbook(Context.getSelectedFile());
var g_headers = ["Capability area", "Capability", "Sub-Capability", "Business Function", "Feature", "#1", "#2", "#3", "#4", "#5", "#6","#7", "#8", "#9", "#10", "#11", "#12", "#13", "#14", "#15", "#16"];
//var g_headers = ["Capability L0", "Capability L1", "Capability L2", "Capability L3", "Capability L4", "Application", "","","","","","","",""];
var g_table = [];
var g_sheetName ="";
var g_columnWidths = [];
g_columnWidths.push([0, 40]);
g_columnWidths.push([1, 40]);
var g_colorGrey = getColorByRGB( 54, 53, 52);
var g_colorLighterGrey = getColorByRGB( 94, 93, 92);
var g_colorLightestGrey = getColorByRGB( 244, 243, 242);
var g_colorGreen = getColorByRGB( 143, 212, 9);


function main(){
    
    //Get Selection
    var oSelModels = ArisData.getSelectedModels();
    var oSelObjs = ArisData.getSelectedObjDefs();
    
    //Objects selected
    if (oSelModels.length ==0){
        var oSelObjs = ArisData.getSelectedObjDefs();
    } 
    else {//Models selected get object definitions from them
        var oSelObjs = [];
        for (var i=0; i<oSelModels.length; i++){
            //var oTempObjs = oSelModels[i].ObjDefListFilter(Constants.OT_FUNC_CLUSTER); //Object type Service type
            var oTempObjs = oSelModels[i].ObjDefListFilter(Constants.OT_IS_FUNC); //Object type Capability
            for (var j=0; j<oTempObjs.length; j++){
                oSelObjs.push(oTempObjs[j]);
                
            }
        }
        oSelObjs = ArisData.Unique(oSelObjs);         //Get unique objects
    }
    
    if (oSelObjs.length == 0){
        return;
    }
    
    oSelObjs= ArisData.sort(oSelObjs, Constants.AT_NAME, g_nLoc);
    var intObjType = oSelObjs[0].TypeNum();
    var strCapabilityLevel = oSelObjs[0].Attribute(g_attrCapabilityLvl, g_nLoc).getValue();
    
    
    
    /*--------------------------------------------------------------------
    TESTING new code using capability level to determine the starting order
    ----------------------------------------------------------------------*/
    
    //Check what capability level it is for the selected object
    if (strCapabilityLevel == "Level 0"){
        for (var c=0; c<oSelObjs.length; c++){
            var oCapabilityL0 = oSelObjs[c];
            var strCapabilityNameL0 = oCapabilityL0.Name(g_nLoc);

            //Capabilities L1
            var oCapabilitiesL1 = oCapabilityL0.getConnectedObjs([Constants.OT_IS_FUNC], Constants.EDGES_OUT);  //Object type capability - get connected L1 capabilities
            oCapabilitiesL1 = ArisData.sort(oCapabilitiesL1, Constants.AT_NAME, g_nLoc);
            
            //Check if no L1 capabilites are found
            if (oCapabilitiesL1.length == 0){
                g_table.push([strCapabilityNameL0, "", "", "", "", ""]);   
            }
           
           //Remove???
           g_table.push([strCapabilityNameL0, "", "", "", "", ""]);
           
           //Iterate through level 1 connected capabilites
            for (var i=0; i<oCapabilitiesL1.length; i++){
                var oCapabilityL1 = oCapabilitiesL1[i];
                var strCapabilityNameL1 = oCapabilityL1.Name(g_nLoc);
                
                //Applications L1
                var oApplicationsL1 = oCapabilityL1.getConnectedObjs([Constants.OT_APPL_SYS_CLS], Constants.EDGES_IN); //Object type application - get connected L1 applications
                oApplicationsL1 = ArisData.sort(oApplicationsL1, Constants.AT_NAME, g_nLoc);
                
                //Capabilities L2
                var oCapabilitiesL2 = oCapabilityL1.getConnectedObjs([Constants.OT_IS_FUNC], Constants.EDGES_OUT);  //Object type capability - get connected L2 capabilities
                oCapabilitiesL2 = ArisData.sort(oCapabilitiesL2, Constants.AT_NAME, g_nLoc);
                
                //Check if no L1 applications and no L2 capabilites are found
                if (oApplicationsL1.length == 0 && oCapabilitiesL2.length == 0){
                    g_table.push([strCapabilityNameL0, strCapabilityNameL1, "", "", "", ""]);   
                }
                
                //Remove???
                g_table.push(["", strCapabilityNameL1, "", "", "", ""]);
                
                //Applications found for L1 
                if (oApplicationsL1.length != 0) {
                    //Setup data row    
                    var oDataRow = [strCapabilityNameL0, strCapabilityNameL1, "", "", ""];
                    var oDataRow = ["", strCapabilityNameL1, "", "", ""];
                    
                    //Applications found for L1 capabilities
                    for (var n=0; n<oApplicationsL1.length; n++){
                        oApplicationL1 = oApplicationsL1[n];
                        strApplicationNameL1 = oApplicationL1.Name(g_nLoc);
                        oDataRow.push(oApplicationL1);  //Adding application data horizontally
                    }
                    //Add L1 capability and application data
                    g_table.push(oDataRow);  
                }
                
                //Iterate through level 2 connected capabilities
                for (var j=0; j<oCapabilitiesL2.length; j++){
                    var oCapabilityL2 = oCapabilitiesL2[j];
                    var strCapabilityNameL2 = oCapabilityL2.Name(g_nLoc);
                    
                    //Applications L2
                    var oApplicationsL2 = oCapabilityL2.getConnectedObjs([Constants.OT_APPL_SYS_CLS], Constants.EDGES_IN); //Object type application - get connected L2 applications
                    oApplicationsL2 = ArisData.sort(oApplicationsL2, Constants.AT_NAME, g_nLoc);
                    
                    //Capabilites L3
                    var oCapabilitiesL3 = oCapabilityL2.getConnectedObjs([Constants.OT_IS_FUNC], Constants.EDGES_OUT);  //Object type capability - get connected L3 capabilities
                    oCapabilitiesL3 = ArisData.sort(oCapabilitiesL3, Constants.AT_NAME, g_nLoc);
                    
                    //Check if no L2 applications and no L3 capabilites are found
                    if(oApplicationsL2.length == 0 && oCapabilitiesL3.length == 0){
                        g_table.push([strCapabilityNameL0, strCapabilityNameL1, strCapabilityNameL2, "", "", ""]);
                    }
                    
                    //Remove???
                    g_table.push(["", "", strCapabilityNameL2, "", "", ""]);
                    
                    //Applications found for L2
                    if (oApplicationsL2.length != 0) {
                        //Setup data row    
                        //var oDataRow = [strCapabilityNameL0, strCapabilityNameL1, strCapabilityNameL2, "", ""];
                        var oDataRow = ["", "", strCapabilityNameL2, "", ""];
                        
                        //Applications found for L2 capabilities
                        for (var n=0; n<oApplicationsL2.length; n++){
                            oApplicationL2 = oApplicationsL2[n];
                            strApplicationNameL2 = oApplicationL2.Name(g_nLoc);
                            oDataRow.push(strApplicationNameL2);  //Adding application data horizontally
                        }
                        //Add L2 capability and application data
                        g_table.push(oDataRow);  
                    }

                    //Iterate through level 3 connected capabilites
                    for(var k=0; k<oCapabilitiesL3.length; k++){
                        var oCapabilityL3 = oCapabilitiesL3[k];
                        var strCapabilityNameL3 = oCapabilityL3.Name(g_nLoc); 
                        
                        //Applications L3
                        var oApplicationsL3 = oCapabilityL3.getConnectedObjs(Constants.OT_APPL_SYS_CLS);  //Object type application - get connected L3 applications
                        oApplicationsL3 = ArisData.sort(oApplicationsL3, Constants.AT_NAME, g_nLoc);
                        
                        //Capabilities L4
                        var oCapabilitiesL4 = oCapabilityL3.getConnectedObjs([Constants.OT_IS_FUNC], Constants.EDGES_OUT);  //Object type capability - get connected L4 capabilities
                        oCapabilitiesL4 = ArisData.sort(oCapabilitiesL4, Constants.AT_NAME, g_nLoc);
                        
                        //Check if no L3 applications and no L4 capabilites are found
                        if (oApplicationsL3.length == 0 && oCapabilitiesL4.length == 0){
                            //g_table.push([strCapabilityNameL0, strCapabilityNameL1, strCapabilityNameL2, strCapabilityNameL3, "", ""]);
                            g_table.push(["", "", "", strCapabilityNameL3, "", ""]);
                        }
                        
                        //Applications found for L3
                        if (oApplicationsL3.length != 0) {
                            //Setup data row    
                            //var oDataRow = [strCapabilityNameL0, strCapabilityNameL1, strCapabilityNameL2, strCapabilityNameL3, ""];
                            var oDataRow = ["", "", "", strCapabilityNameL3, ""];
                            
                            //Applications found for L3 capabilities
                            for (var n=0; n<oApplicationsL3.length; n++){
                                oApplicationL3 = oApplicationsL3[n];
                                strApplicationNameL3 = oApplicationL3.Name(g_nLoc);
                                oDataRow.push(strApplicationNameL3);  //Adding application data horizontally
                            }
                            //Add L3 capability and application data
                            g_table.push(oDataRow);  
                        }

                        //Capabilities L4 have been found
                        if(oCapabilitiesL4.length != 0){
                            //Iterate through L4 capabilites and attempt to get applications
                            for(var n=0; n<oCapabilitiesL4.length; n++){
                                var oCapabilityL4 = oCapabilitiesL4[n];
                                var strCapabilityNameL4 = oCapabilityL4.Name(g_nLoc);
                                
                                //Applications L4
                                var oApplicationsL4 = oCapabilityL4.getConnectedObjs([Constants.OT_APPL_SYS_CLS], Constants.EDGES_IN);  //Object type application - get connected L4 applications
                                oApplicationsL4 = ArisData.sort(oApplicationsL4, Constants.AT_NAME, g_nLoc);
                                
                                //No applications found for L4 capabilites - Add L4 data to table
                                if (oApplicationsL4.length == 0){
                                    //g_table.push([strCapabilityNameL0, strCapabilityNameL1, strCapabilityNameL2, strCapabilityNameL3, strCapabilityNameL4, ""]);
                                    g_table.push(["", "", "", "", strCapabilityNameL4, ""]);
                                    break;
                                }
                                
                                //Setup data row
                                //var oDataRow = [strCapabilityNameL0, strCapabilityNameL1, strCapabilityNameL2, strCapabilityNameL3, strCapabilityNameL4, ""];
                                var oDataRow = ["", "", "", "", strCapabilityNameL4];
                                //Applications found for L4 capabilites
                                for(var m=0; m<oApplicationsL4.length; m++){
                                    oApplicationL4 = oApplicationsL4[m];
                                    strApplicationNameL4 = oApplicationL4.Name(g_nLoc);
                                    oDataRow.push(strApplicationNameL4);  //Adding application data horizontally   
                                } 
                                //Add L4 capability and application data
                                g_table.push(oDataRow);
                            }      
                        }
                    }
                }
            }
        }

    } else if (strCapabilityLevel == "Level 1"){
        //Iterate through level 1 capabilites
        for (var i=0; i<oSelObjs.length; i++){
            var oCapabilityL1 = oSelObjs[i];
            var strCapabilityNameL1 = oCapabilityL1.Name(g_nLoc);
            
            //Applications L1
            var oApplicationsL1 = oCapabilityL1.getConnectedObjs([Constants.OT_APPL_SYS_CLS], Constants.EDGES_IN); //Object type application - get connected L1 applications
            oApplicationsL1 = ArisData.sort(oApplicationsL1, Constants.AT_NAME, g_nLoc);
            
            //Capabilities L2
            var oCapabilitiesL2 = oCapabilityL1.getConnectedObjs([Constants.OT_IS_FUNC], Constants.EDGES_OUT);  //Object type capability - get connected L2 capabilities
            oCapabilitiesL2 = ArisData.sort(oCapabilitiesL2, Constants.AT_NAME, g_nLoc);
            
            //Check if no L1 applications and no L2 capabilites are found
            if(oApplicationsL1.length ==0 && oCapabilitiesL2.length == 0){
                g_table.push(["", strCapabilityNameL1, "", "", "", ""]);
            }
            
            //Remove???
            g_table.push(["", strCapabilityNameL1, "", "", "", ""]);
            
            //Applications found for L1
            if (oApplicationsL1.length != 0) {
                //Setup data row    
                var oDataRow = ["", strCapabilityNameL1, "", "", ""];
                
                //Applications found for L3 capabilities
                for (var n=0; n<oApplicationsL1.length; n++){
                    oApplicationL1 = oApplicationsL1[n];
                    strApplicationNameL1 = oApplicationL1.Name(g_nLoc);
                    oDataRow.push(strApplicationNameL1);  //Adding application data horizontally
                }
                //Add L1 capability and application data
                g_table.push(oDataRow);  
            }

            //Iterate through level 2 connected capabilities
            for (var j=0; j<oCapabilitiesL2.length; j++){
                var oCapabilityL2 = oCapabilitiesL2[j];
                var strCapabilityNameL2 = oCapabilityL2.Name(g_nLoc);
                
                //Applications L2
                var oApplicationsL2 = oCapabilityL2.getConnectedObjs([Constants.OT_APPL_SYS_CLS], Constants.EDGES_IN); //Object type application - get connected L2 applications
                oApplicationsL2 = ArisData.sort(oApplicationsL2, Constants.AT_NAME, g_nLoc);
                
                //Capabilites L3
                var oCapabilitiesL3 = oCapabilityL2.getConnectedObjs([Constants.OT_IS_FUNC], Constants.EDGES_OUT);  //Object type capability - get connected L3 capabilities
                oCapabilitiesL3 = ArisData.sort(oCapabilitiesL3, Constants.AT_NAME, g_nLoc);
                
                //Check if no L2 applications and no L3 capabilites are found
                if(oApplicationsL2.length == 0 && oCapabilitiesL3.length == 0){
                    //g_table.push(["", strCapabilityNameL1, strCapabilityNameL2, "", "", ""]);
                    g_table.push(["", "", strCapabilityNameL2, "", "", ""]);
                }
                //Remove???
                g_table.push(["", "", strCapabilityNameL2, "", "", ""]);
                
                //Applications found for L2
                if (oApplicationsL2.length != 0) {
                    //Setup data row    
                    //var oDataRow = ["", strCapabilityNameL1, strCapabilityNameL2, "", "", ""];
                    var oDataRow = ["", "", strCapabilityNameL2, "", ""];
                    
                    //Applications found for L2 capabilities
                    for (var n=0; n<oApplicationsL2.length; n++){
                        oApplicationL2 = oApplicationsL2[n];
                        strApplicationNameL2 = oApplicationL2.Name(g_nLoc);
                        oDataRow.push(strApplicationNameL2);  //Adding application data horizontally
                    }
                    //Add L2 capability and application data
                    g_table.push(oDataRow);  
                }

                //Iterate through level 3 connected capabilites
                for(var k=0; k<oCapabilitiesL3.length; k++){
                    var oCapabilityL3 = oCapabilitiesL3[k];
                    var strCapabilityNameL3 = oCapabilityL3.Name(g_nLoc); 
                    
                    //Applications L3
                    var oApplicationsL3 = oCapabilityL3.getConnectedObjs(Constants.OT_APPL_SYS_CLS);  //Object type application - get connected L3 applications
                    oApplicationsL3 = ArisData.sort(oApplicationsL3, Constants.AT_NAME, g_nLoc);
                    
                    //Capabilities L4
                    var oCapabilitiesL4 = oCapabilityL3.getConnectedObjs([Constants.OT_IS_FUNC], Constants.EDGES_OUT);  //Object type capability - get connected L4 capabilities
                    oCapabilitiesL4 = ArisData.sort(oCapabilitiesL4, Constants.AT_NAME, g_nLoc);
                                  
                    //Check if no L3 applications and no L4 capabilites are found
                    if (oApplicationsL3.length == 0 && oCapabilitiesL4.length == 0){
                        //g_table.push(["", strCapabilityNameL1, strCapabilityNameL2, strCapabilityNameL3, "", ""]);
                        g_table.push(["", "", "", strCapabilityNameL3, "", ""]);
                    }
                    
                    //Applications found for L3
                    if (oApplicationsL3.length != 0) {
                        //Setup data row    
                        //var oDataRow = ["", strCapabilityNameL1, strCapabilityNameL2, strCapabilityNameL3, "", ""];
                        var oDataRow = ["", "", "", strCapabilityNameL3, ""];
                        
                        //Applications found for L3 capabilities
                        for (var n=0; n<oApplicationsL3.length; n++){
                            oApplicationL3 = oApplicationsL3[n];
                            strApplicationNameL3 = oApplicationL3.Name(g_nLoc);
                            oDataRow.push(strApplicationNameL3);  //Adding application data horizontally
                        }
                        //Add L3 capability and application data
                        g_table.push(oDataRow);  
                    }

                    //Capabilities L4 have been found
                    if(oCapabilitiesL4.length != 0){
                        //Iterate through L4 capabilites and attempt to get applications
                        for(var n=0; n<oCapabilitiesL4.length; n++){
                            var oCapabilityL4 = oCapabilitiesL4[n];
                            var strCapabilityNameL4 = oCapabilityL4.Name(g_nLoc);
                            
                            //Applications L4
                            var oApplicationsL4 = oCapabilityL4.getConnectedObjs([Constants.OT_APPL_SYS_CLS], Constants.EDGES_OUT);  //Object type application - get connected L4 applications
                            oApplicationsL4 = ArisData.sort(oApplicationsL4, Constants.AT_NAME, g_nLoc);
                            
                            //No applications found for L4 capabilites - Add L4 data to table
                            if (oApplicationsL4.length == 0){
                                //g_table.push(["", strCapabilityNameL1, strCapabilityNameL2, strCapabilityNameL3, strCapabilityNameL4, ""]);
                                g_table.push(["", "", "", "", strCapabilityNameL4, ""]);
                                break;
                            }
                            
                            //Setup data row
                            //var oDataRow = ["", strCapabilityNameL1, strCapabilityNameL2, strCapabilityNameL3, strCapabilityNameL4, ""];
                            var oDataRow = ["", "", "", "", strCapabilityNameL4];
                            
                            //Applications found for L4 capabilites
                            for(var m=0; m<oApplicationsL4.length; m++){
                                oApplicationL4 = oApplicationsL4[m];
                                strApplicationNameL4 = oApplicationL4.Name(g_nLoc);
                                oDataRow.push(strApplicationNameL4);  //Adding application data horizontally   
                            } 
                            //Add L4 capability and application data
                            g_table.push(oDataRow);
                            
                        }      
                    }
                }
            }
        }
    } else if (strCapabilityLevel == "Level 2"){
        //Iterate through level 2 capabilities
        for (var j=0; j<oSelObjs.length; j++){
            var oCapabilityL2 = oSelObjs[j];
            var strCapabilityNameL2 = oCapabilityL2.Name(g_nLoc);
            
            //Applications L2
            var oApplicationsL2 = oCapabilityL2.getConnectedObjs([Constants.OT_APPL_SYS_CLS], Constants.EDGES_IN); //Object type application - get connected L2 applications
            oApplicationsL2 = ArisData.sort(oApplicationsL2, Constants.AT_NAME, g_nLoc);
            
            //Capabilities L3
            var oCapabilitiesL3 = oCapabilityL2.getConnectedObjs([Constants.OT_IS_FUNC], Constants.EDGES_OUT);  //Object type capability - get connected L3 capabilities
            oCapabilitiesL3 = ArisData.sort(oCapabilitiesL3, Constants.AT_NAME, g_nLoc);
            
            //Check if no L2 applications and no L3 capabilites are found
            if(oApplicationsL2.length == 0 && oCapabilitiesL3.length == 0){
                g_table.push(["", "", strCapabilityNameL2, "", "", ""]);
            }
            
            //Remove???
            g_table.push(["", "", strCapabilityNameL2, "", "", ""]);
            
            //Applications found for L2
            if (oApplicationsL2.length != 0) {
                //Setup data row    
                //var oDataRow = ["", "", strCapabilityNameL2, "", "", ""];
                var oDataRow = ["", "", strCapabilityNameL2, "", ""];
                
                //Applications found for L2 capabilities
                for (var n=0; n<oApplicationsL2.length; n++){
                    oApplicationL2 = oApplicationsL2[n];
                    strApplicationNameL2 = oApplicationL2.Name(g_nLoc);
                    oDataRow.push(strApplicationNameL2);  //Adding application data horizontally
                }
                //Add L2 capability and application data
                g_table.push(oDataRow);  
            }
            
            //Iterate through level 3 connected capabilites
            for(var k=0; k<oCapabilitiesL3.length; k++){
                var oCapabilityL3 = oCapabilitiesL3[k];
                var strCapabilityNameL3 = oCapabilityL3.Name(g_nLoc);
                
                //Applications L3
                var oApplicationsL3 = oCapabilityL3.getConnectedObjs([Constants.OT_APPL_SYS_CLS], Constants.EDGES_IN); //Object type application - get connected L3 applications
                oApplicationsL3 = ArisData.sort(oApplicationsL3, Constants.AT_NAME, g_nLoc);
                
                //Capabilities L4
                var oCapabilitiesL4 = oCapabilityL3.getConnectedObjs([Constants.OT_IS_FUNC], Constants.EDGES_OUT);  //Object type capability - get connected L4 capabilities
                oCapabilitiesL4 = ArisData.sort(oCapabilitiesL4, Constants.AT_NAME, g_nLoc);
                
                //Check if no L3 applications and no L4 capabilites are found
                if (oApplicationsL3.length == 0 && oCapabilitiesL4.length == 0){
                    //g_table.push(["", "", strCapabilityNameL2, strCapabilityNameL3, "", ""]);  
                    g_table.push(["", "", "", strCapabilityNameL3, "", ""]); 
                }
                
                //Applications found for L3
                if (oApplicationsL3.length != 0) {
                    //Setup data row    
                    //var oDataRow = ["", "", strCapabilityNameL2, strCapabilityNameL3, "", ""];
                    var oDataRow = ["", "", "", strCapabilityNameL3, ""];
                    
                    //Applications found for L3 capabilities
                    for (var n=0; n<oApplicationsL3.length; n++){
                        oApplicationL3 = oApplicationsL3[n];
                        strApplicationNameL3 = oApplicationL3.Name(g_nLoc);
                        oDataRow.push(strApplicationNameL3);  //Adding application data horizontally
                    }
                    //Add L3 capability and application data
                    g_table.push(oDataRow);  
                }
                
                //Capabilities L4 have been found
                if(oCapabilitiesL4.length != 0){
                    //Iterate through L4 capabilites and attempt to get applications
                    for(var n=0; n<oCapabilitiesL4.length; n++){
                        var oCapabilityL4 = oCapabilitiesL4[n];
                        var strCapabilityNameL4 = oCapabilityL4.Name(g_nLoc);
                        
                        //Applications L4
                        var oApplicationsL4 = oCapabilityL4.getConnectedObjs([Constants.OT_APPL_SYS_CLS], Constants.EDGES_IN);  //Object type application - get connected L4 applications
                        oApplicationsL4 = ArisData.sort(oApplicationsL4, Constants.AT_NAME, g_nLoc);
                        
                        //No applications found for L4 capabilites - Add L4 data to table
                        if (oApplicationsL4.length == 0){
                            //g_table.push(["", "", strCapabilityNameL2, strCapabilityNameL3, strCapabilityNameL4, ""]);
                            g_table.push(["", "", "", "", strCapabilityNameL4, ""]);
                            break;
                        } 
                        
                        //Setup data row
                        //var oDataRow = ["", "", strCapabilityNameL2, strCapabilityNameL3, strCapabilityNameL4, ""];
                        var oDataRow = ["", "", "", "", strCapabilityNameL4];
                        
                        //Applications found for L4 capabilites
                        for(var m=0; m<oApplicationsL4.length; m++){
                            oApplicationL4 = oApplicationsL4[m];
                            strApplicationNameL4 = oApplicationL4.Name(g_nLoc);
                            oDataRow.push(strApplicationNameL4);  //Adding application data horizontally   
                        } 
                        //Add L4 capability and application data
                        g_table.push(oDataRow);

                    }      
                }
            }
        }
        
    } else if (strCapabilityLevel == "Level 3"){
        //Iterate through level 3 capabilities
        for (var i=0; i<oSelObjs.length; i++){
            var oCapabilityL3 = oSelObjs[i];
            var strCapabilityNameL3 = oCapabilityL3.Name(g_nLoc);
            
            //Applications L3
            var oApplicationsL3 = oCapabilityL3.getConnectedObjs(Constants.OT_APPL_SYS_CLS);  //Object type application - get connected L3 applications
            oApplicationsL3 = ArisData.sort(oApplicationsL3, Constants.AT_NAME, g_nLoc);
            
            //Capabilities L4
            var oCapabilitiesL4 = oCapabilityL3.getConnectedObjs([Constants.OT_IS_FUNC], Constants.EDGES_OUT);  //Object type capability - get connected L4 capabilities
            oCapabilitiesL4 = ArisData.sort(oCapabilitiesL4, Constants.AT_NAME, g_nLoc);
            
            //Check if no L3 applications and no L4 capabilites are found
            if (oApplicationsL3.length == 0 && oCapabilitiesL4.length == 0){
                g_table.push(["", "", "", strCapabilityNameL3, "", ""]);   
            }
            
            //Applications found for L3
            if (oApplicationsL3.length != 0) {
                //Setup data row    
                //var oDataRow = ["", "", "", strCapabilityNameL3, "", ""];
                var oDataRow = ["", "", "", strCapabilityNameL3, ""];
                
                //Applications found for L3 capabilities
                for (var n=0; n<oApplicationsL3.length; n++){
                    oApplicationL3 = oApplicationsL3[n];
                    strApplicationNameL3 = oApplicationL3.Name(g_nLoc);
                    oDataRow.push(strApplicationNameL3);  //Adding application data horizontally
                }
                //Add L3 capability and application data
                g_table.push(oDataRow);  
            }
            
            //Capabilities L4 have been found
            if(oCapabilitiesL4.length != 0){
                //Iterate through L4 capabilites and attempt to get applications
                for(var n=0; n<oCapabilitiesL4.length; n++){
                    var oCapabilityL4 = oCapabilitiesL4[n];
                    var strCapabilityNameL4 = oCapabilityL4.Name(g_nLoc);
                    
                    //Applications L3
                    var oApplicationsL4 = oCapabilityL4.getConnectedObjs([Constants.OT_APPL_SYS_CLS], Constants.EDGES_IN);  //Object type application - get connected L4 applications
                    oApplicationsL4 = ArisData.sort(oApplicationsL4, Constants.AT_NAME, g_nLoc);
                    
                    //No applications found for L4 capabilites - Add L4 data to table
                    if (oApplicationsL4.length == 0){
                        g_table.push(["", "", "", "", strCapabilityNameL4, ""]);
                        break;
                    }
                    
                    //Setup data row
                    //var oDataRow = ["", "", "", strCapabilityNameL3, strCapabilityNameL4, ""];
                    var oDataRow = ["", "", "", "", strCapabilityNameL4];
                    
                    //Applications found for L4 capabilites
                    for(var m=0; m<oApplicationsL4.length; m++){
                        oApplicationL4 = oApplicationsL4[m];
                        strApplicationNameL4 = oApplicationL4.Name(g_nLoc);
                        oDataRow.push(strApplicationNameL4);  //Adding application data horizontally   
                    } 
                    //Add L4 capability and application data
                    g_table.push(oDataRow);
                }      
            }
        }
        
    } else if (strCapabilityLevel == "Level 4"){
        //Iterate through level 4 capabilities
        for (var i=0; i<oSelObjs.length; i++){
            var oCapabilityL4 = oSelObjs[i];
            var strCapabilityNameL4 = oCapabilityL4.name(g_nLoc);
            
            //Applications L4
            var oApplicationsL4 = oCapabilityL4.getConnectedObjs([Constants.OT_APPL_SYS_CLS], Constants.EDGES_IN);  //Object type application - get connected L4 applications
            oApplicationsL4 = ArisData.sort(oApplicationsL4, Constants.AT_NAME, g_nLoc);
            
            //No applications found for L4 capabilites - Add L4 data to table
            if (oApplicationsL4.length == 0){
                g_table.push(["", "", "", "", strCapabilityNameL4, ""]);
            }
            //Setup data row
            //var oDataRow = ["", "", "", strCapabilityNameL3, strCapabilityNameL4, ""];
            var oDataRow = ["", "", "", "", strCapabilityNameL4];
            
            //Applications found for L4 capabilites
            for(var m=0; m<oApplicationsL4.length; m++){
                oApplicationL4 = oApplicationsL4[m];
                strApplicationNameL4 = oApplicationL4.Name(g_nLoc);
                oDataRow.push(strApplicationNameL4);  //Adding application data horizontally   
            } 
            //Add L4 capability and application data
            g_table.push(oDataRow);
        }
    }

    printInformation(g_table, g_headers, "capability", g_columnWidths);
    g_exc.write();
    
}
main();




/* -------------------------
Print to Excel
------------------------- */
function printInformation(p_table, p_headers, p_sheetName, p_columnWidths) {
    
    var headerFont = g_exc.getFont("Arial", 11, Constants.C_BLACK, true, false, false, false, Constants.XL_FONT_SS_NONE);
    var headerStyle = g_exc.createCellStyle(headerFont, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_LEFT, Constants.ALIGN_CENTER, Constants.C_TRANSPARENT, Constants.C_GREY_80_PERCENT, Constants.FILLTYPE_TRANSPARENT);
    headerStyle.setWrapText(true);
    var normalFont = g_exc.getFont("Calibri", 11, Constants.C_BLACK, false, false, false, false, Constants.XL_FONT_SS_NONE);
    //var normalStyle = g_exc.createCellStyle(normalFont, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_LEFT, Constants.ALIGN_CENTER, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.FILLTYPE_TRANSPARENT);
    
    var normalStyle = g_exc.createCellStyle(normalFont, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_LEFT, Constants.ALIGN_CENTER, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.FILLTYPE_TRANSPARENT);
    normalStyle.setWrapText(true);
    //createCellStyle ( XlsFont p_Font, short p_TopBorder, short p_RightBorder, short p_BottomBorder, short p_LeftBorder, int p_TopBorderColor, int p_RightBorderColor, int p_BottomBorderColor, int p_LeftBorderColor, short p_HorizontalAlignment, short p_VerticalAlignment, int p_BackgroundColor, int p_ForegroundColor, short p_FillPattern )
    var sObjects = g_exc.createSheet(p_sheetName);
    
    var headRow = sObjects.createRow(0);
    for (var i = 0; i < p_headers.length; i++) {
        headRow.createCell(i).setCellValue(p_headers[i]);
    }
    
    
    for (var i = 0; i < p_table.length; i++) {
        var excRow = sObjects.createRow(i + 1);
        for (var j = 0; j < p_table[i].length; j++) {
            writeExcelCell(p_table, excRow, j, j);
            excRow.getCellAt(j).setCellStyle(normalStyle);
        }
    }
    
    //format work sheet
    sObjects.setDefaultColumnWidth(20);
    for (var i = 0; i < p_columnWidths.length; i++) {
        sObjects.setColumnWidth(p_columnWidths[i][0], p_columnWidths[i][1] * 256);
    }
    
    sObjects.setZoom(3, 4);
    sObjects.createFreezePane(0,1);//Header row frozen
    
    //bold Header row
    var headerCells = sObjects.getRowAt(0).getCells();
    for (var i = 0; i < headerCells.length; i++) {
        headerCells[i].setCellStyle(headerStyle);
    }
    
    //mergeSameCells(sObjects, 0);
    //mergeSameCells(sObjects, 1);
    //mergeSameCells(sObjects, 2);
    
    //g_exc.setWriteProtect(true,"pwd","user");
    //sObjects.setProtect(true);
    
    
}
/*--------------------------------------------
Help functions
------------------------------------------*/
function mergeSameCells(p_sheet, p_column, p_referenceColumn) {
    //Merge cells with same value
    if (p_referenceColumn == null){
        p_referenceColumn = 0;
    }
    var lastRowNum = p_sheet.getLastFilledRowNum();
    var lastValue = "";
    var mergeStartRow = 0;
    var mergeEndRow = 0;
    for (var i = 0; i < lastRowNum + 1; i++) {
        var oCell = p_sheet.getCell(i, p_column);
        if (oCell == null) {//there is no value
            if (mergeEndRow > mergeStartRow) {
                p_sheet.addRegion(p_column, p_column, mergeStartRow, mergeEndRow);
            }
            mergeStartRow = i;
            mergeEndRow = i;
            continue;
        }
        var cellValue = oCell.getCellValue();
        
        //Check if there is a change in the previous column
        var changeAtLeftCell = false;
        if (p_column > 0 && i > 0) {//don not check first column or first row
            var oLeftCell = p_sheet.getCell(i, p_referenceColumn);
            var oPrevLeftCell = p_sheet.getCell(i - 1, p_referenceColumn);
            if (oLeftCell != null && oPrevLeftCell != null) {
                var oLeftValue = oLeftCell.getCellValue();
                var oPrevLeftValue = oPrevLeftCell.getCellValue();
                if (!oLeftValue.equals(oPrevLeftValue)) {
                    changeAtLeftCell = true;
                }
            } else {//either cell is null, merge
                changeAtLeftCell = true;
            }
        }
        if (cellValue.equals(lastValue) && !changeAtLeftCell) {//there is no change at this row or the previous
            mergeEndRow = i;
            if (i == lastRowNum && mergeEndRow > mergeStartRow) { //End of table, merge if area longer than 1
                p_sheet.addRegion(p_column, p_column, mergeStartRow, lastRowNum)
            }
        } else {
            
            if (mergeEndRow > mergeStartRow) {//If only one cell in area, do not merge
                p_sheet.addRegion(p_column, p_column, mergeStartRow, mergeEndRow);
            }
            mergeStartRow = i;
            mergeEndRow = i;
        }
        lastValue = cellValue;
    }
}

function changeColumnStyle(p_sheet, p_column, p_style){
    //Merge cells with same value
    var lastRowNum = p_sheet.getLastFilledRowNum();
    var lastValue = "";
    var mergeStartRow = 0;
    var mergeEndRow = 0;
    for (var i=1; i<lastRowNum+1; i++){
        var oCell = p_sheet.getCell(i, p_column);
        if (oCell != null){//there is a value
            oCell.setCellStyle(p_style);
            continue;
        }
    }
}
function writeExcelCell(p_table, p_row, p_tableIndex, p_columnIndex) {
    var rowIndex = p_row.getRowIndex();
    var excCell = p_row.createCell(p_columnIndex);
    var cellData = "";
    if (isArray(p_table[rowIndex - 1][p_tableIndex])) {
        for (i1 = 0; i1 < p_table[rowIndex - 1][p_tableIndex].length; i1++) {
            cellData = cellData + p_table[rowIndex - 1][p_tableIndex][i1] + "; ";
        }
        var nr = cellData.length;
        cellData = cellData.slice(0, nr - 2);
        
    } else {
        cellData = p_table[rowIndex - 1][p_tableIndex];
    }
    if (cellData.length > 4) {
        if (cellData.slice(0, 4) == "http") {
            excCell.setHyperlink(cellData, cellData);
        } else {
            excCell.setCellValue(cellData);
        }
    }
    excCell.setCellValue(cellData);
    
    
}


function isArray(testObject) {
    return testObject && !(testObject.propertyIsEnumerable('length')) && typeof testObject === 'object' && typeof testObject.length === 'number';
}


//Will check if a given item p_item is in a given list p_itemList
function isValidObjectType(p_item, p_itemList){
    var exist = false;    
    for (i=0; i<p_itemList.length; i++){
        if (p_itemList[i] == p_item){
            exist = true;
            return exist;
        }
    }
    return exist;
}
