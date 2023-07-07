// text constants
// dialog text constants
var txtOutputOptionsDialogTitle = getString("TEXT1");
var txtOutputFormat         = getString("TEXT2");
var txtOFModelHier          = getString("TEXT3");
var txtOFModelHierISO       = getString("TEXT4");
var txtSupplement           = getString("TEXT5");

var txtSortOrder            = getString("TEXT6");
var txtSortSymbolTypes      = getString("TEXT7");
var txtSortAlpha            = getString("TEXT8");
var txtSortTopologically    = getString("TEXT9");
var txtSortNumerically      = getString("TEXT10");
var txtLinkLevels           = getString("TEXT11");

var txtModelTypes           = getString("TEXT12");
var txtMTeEPC               = getString("TEXT13");
var txtGraphic              = getString("TEXT19");
var txtFormatGraphic        = getString("TEXT20");

// output text constants
var txtStartPath            = getString("TEXT21");
var txtAssignedModels       = getString("TEXT22");
var txtFunction             = getString("TEXT23");
var txtChapter              = getString("TEXT24");

var txtTopLevel             = getString("TEXT25");
var txtLevel                = getString("TEXT26");
var txtAssignmentOf         = getString("TEXT27");
var txtNull                 = getString("TEXT28");
var txtGroup                = getString("TEXT29");
var txtPicture              = getString("TEXT30");
var txtNormElements         = getString("TEXT31");
var txtNorm                 = getString("TEXT32");
var txtNotNamed             = getString("TEXT33");
var txtRelations            = getString("TEXT34");
var txtInput                = getString("TEXT35");
var txtOutput               = getString("TEXT36");
var txtOrgResp              = getString("TEXT37");
var txtApplSysSupp          = getString("TEXT38");
var txtFrom                 = getString("TEXT39");
var txtLeadsTo              = getString("TEXT40");

var txtFurtherFunctionsWithSameAssignment = getString("TEXT41");

var txtInputObjects         = getString("TEXT42");
var txtOutputObjects        = getString("TEXT43");
var txtOrgElements          = getString("TEXT44");
var txtApplSystems          = getString("TEXT38");

// message box text constants
var txtPleaseNumber         = getString("TEXT45");
var txtNumberToSmall        = getString("TEXT46");
var txtNoModelsSelected     = getString("TEXT47");


var bremote = false; 

// Global variables.
var g_bmodtypedlg = false; 
var g_npicnum = 0; // Numbering model graphic.
var g_nselectedoption = 0; // Selected sort criterion.
var g_nprozfunclevel = 0; // Counter of functions in the active model.
var g_ooutfile = new __holder(null); // Output object.
var g_oworkingarray = null; // List containing the models of the level to be processed right now.
var g_odonemodels = null; // List of models that have already been evaluated.
var g_odonefuncdefs = new __holder(null); // List of functions that have already been evaluated.
var g_odonemodfuncoccs = new __holder(null); // List of function occurrences that have already been topologically evaluated in the model.
var g_oepkfuncobjects = null; // List of functions which are being processed in case of topological processing of the eEPC..
var g_oinputobjdefs = null; // Input objects of the model.
var g_ooutputobjdefs = null; // Output objects of the model.
var g_oorgobjdefs = null; // Organizational objects of the model.
var g_oapplobjdefs = null; // Application system objects of the model.
var g_ssourcearray = __createArray(""); // List telling you by the help of which function the assigned model was found.
var g_sdonefuncident = new Array(); // ID of the evaluated function.
var g_sdfuoccid = __createArray(""); // ID evaluated function.
var g_sdonemodident = __createArray(""); // ID of the evaluated models.
var g_ssortstrings = __createArray(""); // List of strings to be exchanged are put in subroutine SortPosition.
var g_ssourceruleoffunc = __createArray(""); // List containing the predecessor functions of the function in the list g_oEpkFuncObjects.
var g_sselecttion = ""; // Variable containing the type of the model to be evaluated when evaluating topologically.
var g_sdonefuncid = ""; // Contains the string in case of topological evaluation.
var g_nloc = 0; // ID of the language.
var g_bstartlable = false; 
var g_bmodelidententry = false; // Indicator flag for the entering in the list g_sDoneModIdent.
var g_bsamefunction = false; // Indicator flag if the function is already evaluated.
var g_bdonefunc = false; // True, if the function has already been evaluated when evaluating topologically.
var g_bGraphic;// Variable whether the model graphic should be output.
var g_ocurarray = new Array();
var g_objCurrentModel; // Model of the level to be processed right now.
var g_blnIsStartEvent = true;

var g_nDefaultLinkLevels = 3; // default link levels
var g_nmodeltypes = new Array();
g_nmodeltypes[g_nmodeltypes.length] = Constants.MT_EEPC	;
g_nmodeltypes[g_nmodeltypes.length] = Constants.MT_EEPC_MAT;
g_nmodeltypes[g_nmodeltypes.length] = Constants.MT_EEPC_COLUMN;
g_nmodeltypes[g_nmodeltypes.length] = Constants.MT_EEPC_TAB_HORIZONTAL;
g_nmodeltypes[g_nmodeltypes.length] = Constants.MT_EEPC_TAB;
g_nmodeltypes[g_nmodeltypes.length] = Constants.MT_EEPC_ROW;
g_nmodeltypes[g_nmodeltypes.length] = Constants.MT_OFFICE_PROC;
g_nmodeltypes[g_nmodeltypes.length] = Constants.MT_IND_PROC;
g_nmodeltypes[g_nmodeltypes.length] = Constants.MT_PRCS_CHN_DGM;
g_nmodeltypes[g_nmodeltypes.length] = Constants.MT_PCD_MAT;
g_nmodeltypes[g_nmodeltypes.length] = Constants.MT_VAL_ADD_CHN_DGM;
g_nmodeltypes[g_nmodeltypes.length] = Constants.MT_FUNC_ALLOC_DGM;
g_nmodeltypes[g_nmodeltypes.length] = Constants.MT_FUNC_TREE;

g_nmodeltypes[g_nmodeltypes.length] = Constants.MT_BPD_BPMN;
g_nmodeltypes[g_nmodeltypes.length] = Constants.MT_BPMN_ALLOC_DIAGR;
g_nmodeltypes[g_nmodeltypes.length] = Constants.MT_PROCESS_SCHEDULE;
//g_nmodeltypes[g_nmodeltypes.length] = Constants.MT_SIPOC; //Custom 20220201

var g_nSelectedModelTypes;
var g_nOptOutput;
var g_bappendix;

function main()
{
  var sextension = "";   // File ending
  var bchangeext = false;   // Variable for checking whether the file ending was changed.
  var ndummy = 0; 
  var ndepth = new __holder(0);   // Depth of the model level to be evaluated.
  var ntypecounter = 0; 
  var ncheckmsg = 0; 
  var bokmodel = false; 
  var bgeneralmodelok = false; 
  var bnotrightmodel = false; 
  var npos = 0; 
  var nuserdlg = 0;   // Variable for the user dialog box

  // Default settings
  g_bmodelidententry = __toBool(false);

  var binput = false;
  bokmodel = __toBool(false);
  bnotrightmodel = __toBool(false);
  bgeneralmodelok = __toBool(false);
  g_bstartlable = __toBool(true);

  g_nloc = __toLong(Context.getSelectedLanguage());
  g_ooutfile.value = Context.createOutputObject(Context.getSelectedFormat(), Context.getSelectedFile());
  g_ooutfile.value.DefineF("REPORT1", getString("TEXT48"), 14, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0, 0, 0, 0, 0, 0);
  g_ooutfile.value.DefineF("REPORT2", getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0, 0, 0, 0, 0, 0);
  g_ooutfile.value.DefineF("REPORT3", getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0, 0, 0, 0, 0, 0);
  g_ooutfile.value.DefineF("REPORT4", getString("TEXT48"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0, 0, 0, 0, 0, 0);
  g_ooutfile.value.Init(g_nloc);

  var omodels = ArisData.getSelectedModels();
  omodels = __sort(omodels, Constants.AT_NAME, Constants.SORT_NONE, Constants.SORT_NONE, g_nloc);

  g_oworkingarray           = new Array();
  g_odonemodels             = new Array();
  g_odonefuncdefs.value     = new Array();
  g_odonemodfuncoccs.value  = new Array();
  g_oepkfuncobjects         = new Array();
  g_oinputobjdefs           = new Array();
  g_ooutputobjdefs          = new Array();
  g_oorgobjdefs             = new Array();
  g_oapplobjdefs            = new Array();

  g_sdonemodident = __redimArray(g_sdonemodident, new __ArrayInfo(0, 0), "", false); 
  g_bsamefunction = __toBool(false);

  g_npicnum = 1;
  g_ssourceruleoffunc = __redimArray(g_ssourceruleoffunc, new __ArrayInfo(0, 0), "", false); 

  if (omodels.length > 0) {
    var holder_nOptOutput       = new __holder(0);
    var holder_bSupplement      = new __holder(false)
    var holder_nOptSortOrder    = new __holder(0);
    var holder_nLinkLevels      = new __holder(g_nDefaultLinkLevels);
    var holder_aModelTypes      = new __holder(new Array());
    var holder_bGraphic         = new __holder(false);
    
    var nuserdialog = showOutputOptionsDialog(g_ooutfile, holder_nOptOutput, holder_bSupplement, holder_nOptSortOrder, holder_nLinkLevels, holder_aModelTypes, holder_bGraphic);

    if(nuserdialog==0) {
      Context.setScriptError(Constants.ERR_CANCEL);
      return;
    }

    g_nOptOutput            = holder_nOptOutput.value;
    g_nselectedoption       = holder_nOptSortOrder.value;
    g_bappendix             = holder_bSupplement.value;
    ndepth.value            = holder_nLinkLevels.value;
    g_nSelectedModelTypes   = holder_aModelTypes.value;
    g_bGraphic              = holder_bGraphic.value;

    // Beginning of processing the selected options.
    setReportHeaderFooter(g_ooutfile.value, g_nloc, false, false, false);

    evaluate(omodels, ndepth);

    g_ooutfile.value.WriteReport(Context.getSelectedPath(), Context.getSelectedFile());
  }
  else {
    Dialogs.MsgBox(txtNoModelsSelected, vbOKOnly, getString("TEXT49"));
  }
}


function isSelectedModelTypeForAssignment(modelTypeNum)
{
  for(var i=0;i<g_nSelectedModelTypes.length;i++) {
    var mt = g_nmodeltypes[g_nSelectedModelTypes[i]];
    if(modelTypeNum==mt) {
      return true; 
    }
  }

  return false;
}


// --------------------------------------------------------------
// Subprogram AssignedModelsIntoList for inserting the assigned models into the corresponding list
// Parameter
// oCurrentFuncOcc = Current function occurrence.
// nDepth = Depth of the model level to be evaluated.
// nCurrentDepth = Current depth of the model levels.
// nModelLevel =  Model number on the current level.
// sFuncLevel = Number of the current function.
// oAsProzModelList = List of assigned models (of the Processes type).
// sSourceFuncProc =  List containing names of functions to which processes are assigned.

function assignedmodelsintolist(ocurrentfuncocc, ndepth, ncurrentdepth, nmodellevel, sfunclevel, oasprozmodellist, ssourcefuncproc)
{
  var bdonemodel = false;   // Indicator flag if the model has already been evaluated.
  var soutstring = new __holder(""); 
  var bfoundassigment = false;

  // Default settings.
  var bcheck = false; 
  var omodelsoffuncass  = ocurrentfuncocc.value.ObjDef().AssignedModels(); // List of models which are assigned to the current function.
  omodelsoffuncass      = __sort(omodelsoffuncass, Constants.AT_NAME, Constants.SORT_NONE, Constants.SORT_NONE, g_nloc);

  if (omodelsoffuncass.length > 0) {
    for (var j = 0 ; j < (omodelsoffuncass.length - 1)+1 ; j++ ) {
      var ocurrentmodeloffuncass = omodelsoffuncass[j];
      bcheck = __toBool(checkassignedmodel(ocurrentmodeloffuncass, oasprozmodellist, ndepth, ncurrentdepth, nmodellevel, soutstring));
      if (isSelectedModelTypeForAssignment(ocurrentmodeloffuncass.OrgModelTypeNum())) {     // TANR 216764
        switch(ocurrentmodeloffuncass.OrgModelTypeNum()) {      // TANR 216764
          case Constants.MT_EEPC:
          case Constants.MT_EEPC_MAT:
          case Constants.MT_IND_PROC:
          case Constants.MT_OFFICE_PROC:
          case Constants.MT_PRCS_CHN_DGM:
          case Constants.MT_PCD_MAT:
          case Constants.MT_UML_ACTIVITY_DGM:
          case Constants.MT_EEPC_COLUMN:
          case Constants.MT_VAL_ADD_CHN_DGM:
          case Constants.MT_EEPC_ROW:
          case Constants.MT_EEPC_TAB:
          case Constants.MT_FUNC_ALLOC_DGM:
          case Constants.MT_EEPC_TAB_HORIZONTAL:

            if (bcheck == false) {
              oasprozmodellist.value[__arrlength(oasprozmodellist)] = ocurrentmodeloffuncass;
              ssourcefuncproc.value = __redimArray(ssourcefuncproc.value, new __ArrayInfo(0, (oasprozmodellist.value.length - 1)), "", true); 
              ssourcefuncproc.value[__getArrayIndex(ssourcefuncproc.value,1,__getUBound(ssourcefuncproc.value, 1))] = ((((((vbStr(ncurrentdepth.value) + ".") + vbStr(nmodellevel.value)) + ".") + sfunclevel.value) + vbChr(9)) + ocurrentfuncocc.value.ObjDef().Name(g_nloc));
              bfoundassigment = true;
            }
          break;

          default:
            // The assigned models of the list are no processes, thus the hierarchy level will not be increased.
            if (bcheck == false) {
              g_oworkingarray[__arrlength(g_oworkingarray)] = ocurrentmodeloffuncass;
              g_ssourcearray = __redimArray(g_ssourcearray, new __ArrayInfo(0, (g_oworkingarray.length - 1)), "", true); 
              g_ssourcearray[__getArrayIndex(g_ssourcearray,1,__getUBound(g_ssourcearray, 1))] = ((((((vbStr(ncurrentdepth.value) + ".") + vbStr(nmodellevel.value)) + ".") + sfunclevel.value) + vbChr(9)) + ocurrentfuncocc.value.ObjDef().Name(g_nloc));
              bfoundassigment = true;
            }
        }
      }
    }
  }
}



// --------------------------------------------------------------
// Subprogram CheckAssignedModel Checks whether the assigned model is taken into account when the output is executed.
// Parameter
// oCurrentAssModel = Current assigned model.
// oAsProzModelList = List of assigned models (of the Processes type).
// nCurrentDepth = Current depth of the model levels.
// nModelLevel =  Model number on the current level.
// sIdOfAsMod = ID of the assigned model.

function checkassignedmodel(ocurrentassmodel, oasprozmodellist, ndepth, ncurrentdepth, nmodellevel, sidofasmod)
{
  var __functionResult = false;

  sidofasmod.value = __toString(ocurrentassmodel.Name(g_nloc));
  // Checks whether model is already evaluated.
  if (g_odonemodels.length > 0) {
    for (var i = 0 ; i < (g_odonemodels.length - 1)+1 ; i++ ){
      if (g_odonemodels[i].IsEqual(ocurrentassmodel) == true) {
        __functionResult = true;
        sidofasmod.value = __toString(g_sdonemodident[__getArrayIndex(g_sdonemodident,1,i)]);
        return __functionResult;
      }
    }
  }

  // Current level is registered for evaluation.
  if (g_odonemodels.length > 0) {
    for (var i = (nmodellevel.value - 1) ; i < (g_oworkingarray.length - 1)+1 ; i++ ){
      if (g_oworkingarray[i].IsEqual(ocurrentassmodel) == true) {
        __functionResult = true;
        sidofasmod.value = __toString(((((__toString(ncurrentdepth.value) + ".") + __toString((i + 1))) + " ") + ocurrentassmodel.Name(g_nloc)));
        return __functionResult;
      }
    }
  }

  // Next level is registered for evaluation.
  if (oasprozmodellist.value.length > 0) {
    for (var i = 0 ; i < (oasprozmodellist.value.length - 1)+1 ; i++ ){
      if (oasprozmodellist.value[i].IsEqual(ocurrentassmodel) == true) {
        __functionResult = true;
        if ((ncurrentdepth.value + 1) <= ndepth.value) {
          sidofasmod.value = __toString(((((__toString((ncurrentdepth.value + 1)) + ".") + __toString((i + 1))) + " ") + ocurrentassmodel.Name(g_nloc)));
        } else {
          sidofasmod.value = __toString(ocurrentassmodel.Name(g_nloc));
        }

        return __functionResult;
      }
    }
  }

  return __functionResult;
}



// --------------------------------------------------------------
// Subprogram CheckObj for checking whether the current object has already been written into a reference list.
// Parameter
// oCurrentObj = Current object (Occ/Def).
// sIDOfEvaluateObj = Designation of the object from the corresponding list if it occurs in the reference list.
// nCurrentDepth = Current depth of the model levels.
// nModelLevel = Model number on the current level.
// sFuncLevel = Number of the current function.
// bEntry = True , and entering into the list.
// bID = True if ID of the element should be returned.
// oReferenceList = List of reference objects.
// sIDList = List of designations.
// sName = Name of the object to be checked.

function checkobj(ocurrentobj, sidofevaluateobj, ncurrentdepth, nmodellevel, sfunclevel, bentry, bid, oreferencelist, sidlist, sname)
{
  var __functionResult = false;

  if (oreferencelist.value.length > 0) {
    for (var i = 0 ; i < (oreferencelist.value.length - 1)+1 ; i++ ){
      if (ocurrentobj.value.IsEqual(oreferencelist.value[i])) {
        if (bid == true) {
          sidofevaluateobj.value = __toString(sidlist[i]);
        }
        __functionResult = true;
        break;
      }
      else {
        if (bid == true) {
          sidofevaluateobj.value = __toString("");
        }
      }
    }
  }

  if (__functionResult == false && bentry == true) {
    if (oreferencelist.value.length == 0) {
        sidlist.length = 0;                      // MWZ, 04.07.06, Call-ID 122680
//      sidlist = __redimArray(sidlist, new __ArrayInfo(0, 0), "", false); 
//    } else {
//      sidlist = __redimArray(sidlist, new __ArrayInfo(0, (__getUBound(sidlist, 1) + 1)), "", true); 
    }
//    oreferencelist.value[__arrlength(oreferencelist)] = ocurrentobj.value;
    oreferencelist.value[oreferencelist.value.length] = ocurrentobj.value;
    sidlist[sidlist.length] = ((((((vbStr(ncurrentdepth.value) + ".") + vbStr(nmodellevel.value)) + ".") + sfunclevel.value) + vbChr(9)) + sname);
  }

  return __functionResult;
}



// --------------------------------------------------------------
// Subprogram eEPKOut for the topological output of the functions in the model (eEPC) after the control flow.
// oCurrentModel = Current model.
// nDepth = Depth of the model level to be evaluated.
// nCurrentDepth = Current depth of the model levels.
// nModelLevel = Number or quantity of models that have already been output.
// oAsProzModelList = List of assigned models (of the Processes type).
// sSourceFuncProc =  List containing names of functions to which processes are assigned.
// oRootList = List of root objects.

function eepkout(ocurrentmodel, ndepth, ncurrentdepth, nmodellevel, oasprozmodellist, ssourcefuncproc, orootlist)
{
  var ndummy = new __holder(0);
  var sdummy = new __holder("");

  var ocurrentoccfunc = new __holder(null); 
  var sfunclevel = ""; 
  var bfirst = false; 

  // Default settings
  g_nprozfunclevel = __toInteger(0);

  if (orootlist.length > 0) {
    for (var i = 0 ; i < (orootlist.length - 1)+1 ; i++ ){
      ocurrentoccfunc.value = orootlist[i];
      bfirst = __toBool(true);
      g_sselecttion = txtMTeEPC;
      //g_ooutfile.value.OutputLn(txtStartPath, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
      g_ooutfile.value.OutputLn("", getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
      outofepkfunc(ocurrentmodel, ocurrentoccfunc, ndepth, ncurrentdepth, nmodellevel, oasprozmodellist, ssourcefuncproc, "", false);
      if (g_oepkfuncobjects.length > 0) {
        var j = 0;
        while (j < g_oepkfuncobjects.length) {
          var bdonefunc = checkobj(new __holder(g_oepkfuncobjects[j]), sdummy, ndummy, ndummy, sdummy, false, false, g_odonemodfuncoccs, g_sdfuoccid, g_oepkfuncobjects[j].ObjDef().Name(g_nloc));
          if (! (bdonefunc)) {
            //g_ooutfile.value.OutputLn(txtStartPath, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
            //g_ooutfile.value.OutputLn("", getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
            (ocurrentmodel, new __holder(g_oepkfuncobjects[j]), ndepth, ncurrentdepth, nmodellevel, oasprozmodellist, ssourcefuncproc, __toString(j), true);
          }
          j = (j + 1);
        }

        // Delete List g_oEpkFuncObjects.
        while (g_oepkfuncobjects.length > 0) {
          g_oepkfuncobjects = __delete(g_oepkfuncobjects, 0);
        }
      }
      g_ssourceruleoffunc = __redimArray(g_ssourceruleoffunc, new __ArrayInfo(0, 0), "", false); 
    }
  }
}



// ----------------------------------------------------------------
// 'Evaluate' for calling the subprogram; 'Manufacture' for recursive processing.
// oModels = List of the selected models.
// nDepth = Depth of the model level to be evaluated.

function evaluate(omodels, ndepth)
{
  g_ssourcearray = __redimArray(g_ssourcearray, new __ArrayInfo(0, (omodels.length - 1)), "", false); 
  // Default settings
  if (ndepth.value > 0) {
    // List of selected models.
    for (var i = 0 ; i < (omodels.length - 1)+1 ; i++ ){
      g_bstartlable = __toBool(true);
      var ocurrentmodel = omodels[i];
      g_oworkingarray[__arrlength(g_oworkingarray)] = ocurrentmodel;
      // Inserts models in g_oWorkingArray for processing in 'Manufacture'.
      g_ssourcearray[__getArrayIndex(g_ssourcearray,1,i)] = getString("TEXT28");
    }
    manufacture(ndepth, new __holder(1));
  }
}



// --------------------------------------------------------------
// Subprogram FindNextFunc for determining the next function.
// oCurrentModel = Current model.
// oCurrentObjOcc = Current ObjOcc by which the next function will be found.
// oTargetFuncObjectList = List of functions which were found over the connections.

function findnextfunc(ocurrentmodel, ocurrentobjocc, otargetfuncobjectlist)
{
  var onextobjoccs = null;   // List of following elements.
  var onextcurrentobjocc = new __holder(null);   // Current following element.
  var bcheck = false; 
  var sdummy = new __holder(""); 
  var ndummy = new __holder(0); 
  var sdummylist = __createArray(""); 

  onextobjoccs = ocurrentmodel.value.GetSuccNodes(ocurrentobjocc);
  bcheck = __toBool(true);
  if (onextobjoccs.length > 0) {
    for (var i = 0 ; i < (onextobjoccs.length - 1)+1 ; i++ ){
      onextcurrentobjocc.value = onextobjoccs[i];
      switch(onextcurrentobjocc.value.ObjDef().TypeNum()) {
        case Constants.OT_FUNC:
          bcheck = __toBool(checkobj(onextcurrentobjocc, sdummy, ndummy, ndummy, sdummy, false, false, otargetfuncobjectlist, sdummylist, ""));
          if (bcheck == false) {
            otargetfuncobjectlist.value[__arrlength(otargetfuncobjectlist)] = onextcurrentobjocc.value;
          }
        break;

        default:
          findnextfunc(ocurrentmodel, onextcurrentobjocc.value, otargetfuncobjectlist);
        }
    }
  }
}



function getstartnodes(ocurrentmodel)
{
  var ostartobjoccs = null; 
  if (ocurrentmodel.OrgModelTypeNum() == Constants.MT_VAL_ADD_CHN_DGM) {        // TANR 216764
    ostartobjoccs = new Array();
    var ofuncoccs = ocurrentmodel.ObjOccListFilter(Constants.OT_FUNC);
    for (var i = 0 ; i < (ofuncoccs.length - 1)+1 ; i++ ){
      var ofuncocc = ofuncoccs[i];
      if (ofuncocc.InDegree(Constants.EDGES_STRUCTURE) == 0) {
        ostartobjoccs[ostartobjoccs.length] = ofuncocc;
      }
    }
  } else {
    ostartobjoccs = ocurrentmodel.StartNodeList();
  }
  ostartobjoccs = __sort(ostartobjoccs, Constants.SORT_GEOMETRIC, Constants.SORT_NONE, Constants.SORT_NONE, g_nloc);
  return ostartobjoccs;
} 



// --------------------------------------------------------------
// FindRootFunc For determining the root functions of the current model.
// oCurrentModel = Current model.
// oRootList = List of root objects.

function findrootfunc(ocurrentmodel, orootlist)
{
  var bcheck = false; 
  var nnumoutcxn = 0; 
  var nincxn = 0; 

  orootlist.value = new Array();
  ocurrentmodel.value.BuildGraph(true);
//  var ostartobjoccs = ocurrentmodel.value.StartNodeList(); // List of start objects.
  var ostartobjoccs = getstartnodes(ocurrentmodel.value);
  if (ostartobjoccs.length > 0) {
    // Elements that are not of the function or event type will be removed from the list.
    while (bcheck == false) {
      bcheck = __toBool(true);
      for (var i = 0 ; i < (ostartobjoccs.length - 1)+1 ; i++ ){
        nnumoutcxn = __toLong(ostartobjoccs[i].OutDegree(Constants.EDGES_STRUCTURE));
        nincxn = __toLong(ostartobjoccs[i].InDegree(Constants.EDGES_STRUCTURE));
        if (! (ostartobjoccs[i].ObjDef().TypeNum() == Constants.OT_FUNC) && ! (ostartobjoccs[i].ObjDef().TypeNum() == Constants.OT_EVT) || (nnumoutcxn == 0 && nincxn != 0)) {
          ostartobjoccs = __delete(ostartobjoccs, ostartobjoccs[i]);
          bcheck = __toBool(false);
          break;
        }
      }
    }
  }

  if (ostartobjoccs.length > 0) {
    for (var i = 0 ; i < (ostartobjoccs.length - 1)+1 ; i++ ){
      var ocurrentstartobjocc = ostartobjoccs[i];
      if (ocurrentstartobjocc.ObjDef().TypeNum() == Constants.OT_FUNC) {
        orootlist.value[__arrlength(orootlist)] = ocurrentstartobjocc;
      } else {
        findnextfunc(ocurrentmodel, ocurrentstartobjocc, orootlist);
      }
    }
  }
}



// --------------------------------------------------------------
// FunctionTreeOut for the topological output of the functions in the model (function tree) after control flow.
// oCurrentModel = Current model.
// nDepth = Depth of the model level to be evaluated.
// nCurrentDepth = Current depth of the model levels.
// nModelLevel = Number or quantity of models that have already been output.
// oAsProzModelList = List of assigned models (of the Processes type).
// sSourceFuncProc =  List containing names of functions to which processes are assigned.
// oRootList = List of root objects.

function functiontreeout(ocurrentmodel, ndepth, ncurrentdepth, nmodellevel, oasprozmodellist, ssourcefuncproc, rootlist)
{
  var ocurrentoccfunc = new __holder(null); 
  var sfunclevel = new __holder(""); 
  var bfirst = new __holder(false); 

  if (rootlist.length > 0) {
    for (var i = 0 ; i < (rootlist.length - 1)+1 ; i++ ){
      ocurrentoccfunc.value = rootlist[i];
      bfirst.value = __toBool(true);
      sfunclevel.value = __toString(vbStr((i + 1)));
      outofthetreefunc(ocurrentmodel, ocurrentoccfunc, ndepth, ncurrentdepth, nmodellevel, sfunclevel, oasprozmodellist, ssourcefuncproc, bfirst, new __holder(""), new __holder(true));
      ocurrentoccfunc.value = null;
    }
  }
}



// --------------------------------------------------------------
// Subprogram 'Manufacture' for the recursive workout of the hierarchy levels.
// nDepth = Depth of the model level to be evaluated.
// nCurrentDepth = Current depth of the model levels.

function manufacture(ndepth, ncurrentdepth)
{
  if(g_nOptOutput == 0) {
    manufacture_std(ndepth, ncurrentdepth);
  } else if(g_nOptOutput == 1) {
    manufacture_iso(ndepth, ncurrentdepth);
  }
}

function manufacture_std(ndepth, ncurrentdepth)
{
  var oasprozmodellist = new __holder(null);   // List of the assigned models (of the processes type).
  var ocurrentmodel = new __holder(null); 
  var ssourcefuncproc = new __holder(__createArray(""));   // List containing names of functions to which processes are assigned.
  var osourcefuncproc = null;   // List of function occurrences to which processes are assigned.
  var ssourcefunction = new __holder("");   // ID of the function to which the current model was assigned.
  var bfirstmodel = false; 

  // Default settings.
  oasprozmodellist.value = new Array();
  var nsourcefuncprocnumbers = 0;

  if (ncurrentdepth.value <= ndepth.value +1) {
    var i = 0;
    while (i < g_oworkingarray.length) {
      if (i == 0) {
        bfirstmodel = __toBool(true);
      } else {
        bfirstmodel = __toBool(false);
      }

      ssourcefunction.value = __toString(g_ssourcearray[__getArrayIndex(g_ssourcearray,1,i)]);
      ocurrentmodel.value = g_oworkingarray[i];
      g_objCurrentModel = ocurrentmodel.value;
      if(i!=0) {
        g_ooutfile.value.OutputField(Constants.FIELD_NEWPAGE, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_CENTER);
      }
      g_odonemodfuncoccs.value = new Array();
      g_oepkfuncobjects = new Array();

      g_sdfuoccid = __redimArray(g_sdfuoccid, new __ArrayInfo(0, 0), "", false); 
      g_ssortstrings = __redimArray(g_ssortstrings, new __ArrayInfo(0, 0), "", false); 

      outputmodeldata(ocurrentmodel, ncurrentdepth, new __holder((i + 1)), ssourcefunction, bfirstmodel);
      g_odonemodels[__arrlength(g_odonemodels)] = ocurrentmodel.value;
      if (g_bmodelidententry == false) {
        g_bmodelidententry = __toBool(true);
      } else {
        g_sdonemodident = __redimArray(g_sdonemodident, new __ArrayInfo(0, (__getUBound(g_sdonemodident, 1) + 1)), "", true); 
      }

      g_sdonemodident[__getArrayIndex(g_sdonemodident,1,__getUBound(g_sdonemodident, 1))] = (((vbStr(ncurrentdepth.value) + ".") + vbStr((i + 1))) + ocurrentmodel.value.Name(g_nloc));
      
      outputfuncofmodels(ocurrentmodel, ndepth, ncurrentdepth, new __holder((i + 1)), oasprozmodellist, ssourcefuncproc);
      g_blnIsStartEvent = true;
      
      if(ocurrentmodel.value.TypeNum() != Constants.MT_VAL_ADD_CHN_DGM)
        writeEndEvents(ocurrentmodel.value);
    
      i = (i + 1);
      ocurrentmodel.value = null;
    }

    // Models of the next hierarchy level in the list (g_oWorkingArray ;g_oWorkingArray must be deleted beforehand).
      g_oworkingarray = new Array();        // MWZ, 04.07.06, Call-ID 122680
//    g_odonemodfuncoccs = new Array();

    if (oasprozmodellist.value.length > 0) {
      g_ssourcearray = __redimArray(g_ssourcearray, new __ArrayInfo(0, (oasprozmodellist.value.length - 1)), "", false); 
      for ( i = 0 ; i < (oasprozmodellist.value.length - 1)+1 ; i++ ){
        g_oworkingarray[__arrlength(g_oworkingarray)] = oasprozmodellist.value[i];
        g_ssourcearray = __redimArray(g_ssourcearray, new __ArrayInfo(0, (__getUBound(g_ssourcearray, 1) + 1)), "", true); 
        g_ssourcearray[__getArrayIndex(g_ssourcearray,1,i)] = ssourcefuncproc.value[__getArrayIndex(ssourcefuncproc.value,1,i)];
      }

      // Raise hierarchy level and output models.
      manufacture(ndepth, new __holder((ncurrentdepth.value + 1)));
    }
  }
}


// --------------------------------------------------------------
// Subprogram 'Manufacture' for the recursive workout of the hierarchy levels.
function manufacture_iso(ndepth, ncurrentdepth)
{
  var oasprozmodellist = new __holder(null); 
  var ocurrentmodel = new __holder(null); 
  var ssourcefuncproc = new __holder(__createArray("")); 
  var osourcefuncproc = null;   // List of function occurrences to which processes are assigned.
  var ssourcefunction = new __holder("");   // ID of the function to which the current model was assigned.
  var bfirstmodel = false; 

  var i = 0; 

  // Default settings.
  oasprozmodellist.value = new Array();
  var nsourcefuncprocnumbers = undefined; 
  nsourcefuncprocnumbers = 0;
  g_sdonemodident = __redimArray(g_sdonemodident, new __ArrayInfo(0, 0), "", false); 

  g_odonemodels = new Array();
  
  g_ocurarray = g_oworkingarray;

  if (ncurrentdepth.value <= ndepth.value +1) {
    i = __toLong(0);
    while (i < g_ocurarray.length) {
      if(i!=0) {
        g_ooutfile.value.OutputField(Constants.FIELD_NEWPAGE, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_CENTER);
      }
      if (i == 0) {
        bfirstmodel = __toBool(true);
      } else {
        bfirstmodel = __toBool(false);
      }

      ssourcefunction.value = __toString(g_ssourcearray[__getArrayIndex(g_ssourcearray,1,i)]);
      ocurrentmodel.value = g_ocurarray[i];

      g_odonemodfuncoccs.value = new Array();

      g_sdfuoccid = __redimArray(g_sdfuoccid, new __ArrayInfo(0, 0), "", false); 

      g_oinputobjdefs = new Array();
      g_ooutputobjdefs = new Array();
      g_oorgobjdefs = new Array();
      g_oapplobjdefs = new Array();
      g_oepkfuncobjects = new Array();

      g_ssortstrings = __redimArray(g_ssortstrings, new __ArrayInfo(0, 0), "", false); 
      outputmodeldata(ocurrentmodel, ncurrentdepth, new __holder((i + 1)), ssourcefunction, bfirstmodel);
      g_odonemodels[__arrlength(g_odonemodels)] = ocurrentmodel.value;
      if (! (g_bmodelidententry)) {
        g_bmodelidententry = __toBool(true);
      } else {
        g_sdonemodident = __redimArray(g_sdonemodident, new __ArrayInfo(0, (__getUBound(g_sdonemodident, 1) + 1)), "", true); 
      }

      g_sdonemodident[__getArrayIndex(g_sdonemodident,1,__getUBound(g_sdonemodident, 1))] = (((vbStr(ncurrentdepth.value) + ".") + vbStr((i + 1))) + ocurrentmodel.value.Name(g_nloc));
      g_ooutfile.value.OutputLn(getString("TEXT50"), getString("TEXT48"), 14, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0);
      outputfuncofmodels(ocurrentmodel, ndepth, ncurrentdepth, new __holder((i + 1)), oasprozmodellist, ssourcefuncproc);
      if (g_bappendix) {
        outobjofmodels(g_oinputobjdefs, txtInputObjects + ": ");
        outobjofmodels(g_ooutputobjdefs, txtOutputObjects + ": ");
        outobjofmodels(g_oorgobjdefs, txtOrgElements + ": ");
        outobjofmodels(g_oapplobjdefs, txtApplSystems + ": ");
      }
      i++;
      ocurrentmodel.value = null;
    }

    // Models of the next hierarchy level in the list (g_oCurArray ;g_oCurArray must be deleted beforehand)
    g_oworkingarray = new Array();

    if (oasprozmodellist.value.length > 0) {
      g_ssourcearray = __redimArray(g_ssourcearray, new __ArrayInfo(0, (oasprozmodellist.value.length - 1)), "", false); 
      for ( i = 0 ; i < (oasprozmodellist.value.length - 1)+1 ; i++ ){
        g_oworkingarray[__arrlength(g_oworkingarray)] = oasprozmodellist.value[i];
        g_ssourcearray = __redimArray(g_ssourcearray, new __ArrayInfo(0, (__getUBound(g_ssourcearray, 1) + 1)), "", true); 
        g_ssourcearray[__getArrayIndex(g_ssourcearray,1,i)] = ssourcefuncproc.value[__getArrayIndex(ssourcefuncproc.value,1,i)];
      }

      // Raise hierarchy level and output models.
      manufacture(ndepth, new __holder((ncurrentdepth.value + 1)));
    }
  }
  oasprozmodellist.value = null;
}


// --------------------------------------------------------------
// Subprogram OutFuncData for outputting the functions' data.
// oCurrentFuncOcc = Current function occurrence.
// nDepth = Depth of the model level to be evaluated.
// nCurrentDepth = Current depth of the model levels.
// nModelLevel =  Model number on the current level.
// sFuncLevel = Number of the current function.
// oAsProzModelList = List of assigned models (of the Processes type).
// sSourceFuncProc =  List containing names of functions to which processes are assigned.
// sBough = Designation of the branch when outputting function trees.

function outfuncdata(ocurrentfuncocc, ndepth, ncurrentdepth, nmodellevel, sfunclevel, oasprozmodellist, ssourcefuncproc, sbough)
{
  var ocxnocclist           = new __holder(null); 
  var sidentificationofassignedmodels = __createArray(""); 
  var sidofevaluatefun      = new __holder(""); 
  var ssearchleftstring     = new __holder(""); 
  var ssearchrightstring    = new __holder(""); 

  // Default settings
  var iidentificationofassignedmodelsnumbers = 0; 
  var bdonefunc2    = __toBool(checkobj(ocurrentfuncocc, sidofevaluatefun, ncurrentdepth, nmodellevel, sfunclevel, true, true, g_odonemodfuncoccs, g_sdfuoccid, ocurrentfuncocc.value.ObjDef().Name(g_nloc)));
  var bdonefunc     = __toBool(checkobj(new __holder(ocurrentfuncocc.value.ObjDef()), sidofevaluatefun, ncurrentdepth, nmodellevel, sfunclevel, true, true, g_odonefuncdefs, g_sdonefuncident, ocurrentfuncocc.value.ObjDef().Name(g_nloc)));

  if (! (bdonefunc2)) {
    var ssearchchar = __toString(vbChr(9));
    ssearchleftstring.value = __toString("");
    ssearchrightstring.value = __toString("");
    // Assignments of the function are put in the corresponding list.
    var bcallfunction = __toBool(false);
    if (ncurrentdepth.value <= ndepth.value) 
    {
      if (! (ocurrentfuncocc.value.SymbolNum() == Constants.ST_PRCS_IF)) {
        assignedmodelsintolist(ocurrentfuncocc, ndepth, ncurrentdepth, nmodellevel, sfunclevel, oasprozmodellist, ssourcefuncproc);
      }
    }

    if (bdonefunc == false) 
    {
      var oFunctionDef = ocurrentfuncocc.value.ObjDef();
      var bIsControlActivity = oFunctionDef.Attribute(Constants.AT_USER_ATTR_BOOL_100,g_nloc).GetValue(false);   
    
      //if(bIsControlActivity != "True")
          _writePrecedingEvents(oFunctionDef);
      
      //g_ooutfile.value.OutputLnF(((((((vbStr(ncurrentdepth.value) + ".") + vbStr(nmodellevel.value)) + ".") + sfunclevel.value) + " ") + ocurrentfuncocc.value.ObjDef().Name(g_nloc)), "REPORT3");
      g_ooutfile.value.OutputLnF("", "REPORT3");
      g_ooutfile.value.BeginTable(100, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
      //if (g_sselecttion == getString("TEXT13") && ! (sbough.value == "")) {
      //  g_ooutfile.value.TableRow();
      //  g_ooutfile.value.TableCell(txtFrom+" ", 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
      //  g_ooutfile.value.TableCell(sbough.value, 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
      //}

      //if (g_bdonefunc == true) {
      //  stringcut(g_sdonefuncid, ssearchleftstring, ssearchrightstring, ssearchchar);
      //  g_ooutfile.value.TableRow();
      //  g_ooutfile.value.TableCell(txtLeadsTo+" ", 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
      //  g_ooutfile.value.TableCell(((((txtFunction + "  " + ssearchrightstring.value) + getString("TEXT51")) + ssearchleftstring.value) + ")"), 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
      //}

      var oattributes = ocurrentfuncocc.value.ObjDef().AttrList(g_nloc);
      if (oattributes.length > 0)
      {
          if(ocurrentfuncocc.value.Model().TypeNum() != Constants.MT_VAL_ADD_CHN_DGM)
             _outputFunctionAttributes(oFunctionDef);
            //outofattributes(oattributes);
      }

      ocxnocclist.value = ocurrentfuncocc.value.CxnOccList();
      if (ocxnocclist.value.length > 0)
      {
          if(ocurrentfuncocc.value.Model().TypeNum() != Constants.MT_VAL_ADD_CHN_DGM)
            _outputRelationships(oFunctionDef);
            //outofrelationships(ocxnocclist, ocurrentfuncocc);
      }

      // Assignments
      //outofassignedmodels(ocurrentfuncocc, oasprozmodellist, ndepth, ncurrentdepth, nmodellevel);
      g_ooutfile.value.EndTable("", 100, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    }
    else
    {
      /*stringcut(sidofevaluatefun.value, ssearchleftstring, ssearchrightstring, ssearchchar);
      g_ooutfile.value.OutputLnF(((((((((vbStr(ncurrentdepth.value) + ".") + vbStr(nmodellevel.value)) + ".") + sfunclevel.value) + "  ") + ssearchrightstring.value) + " " + getString("TEXT52")) + " " + ssearchleftstring.value), "REPORT3");
      // Assignments
      g_ooutfile.value.BeginTable(100, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
      ocxnocclist.value = ocurrentfuncocc.value.CxnOccList();
      if (ocxnocclist.value.length > 0)
      {
          if(ocurrentfuncocc.value.Model().TypeNum() != Constants.MT_VAL_ADD_CHN_DGM)
            outofrelationships(ocxnocclist, ocurrentfuncocc);
      }
      g_ooutfile.value.EndTable("", 100, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);*/
    }
  }
}



// --------------------------------------------------------------
// Subroutine OutNum for outputting the functions contained in the model which are sorted on the basis of the numbering of the attribute type 1.
// oCurrentModel = Current model.
// nDepth = Depth of the model level to be evaluated.
// nCurrentDepth = Current depth of the model levels.
// nModelLevel = Number or quantity of models that have already been output.
// oAsProzModelList = List of assigned models (of the Processes type).
// sSourceFuncProc =  List containing names of functions to which processes are assigned.

function outnum(ocurrentmodel, ndepth, ncurrentdepth, nmodellevel, oasprozmodellist, ssourcefuncproc)
{
  var ocurrentfuncocc = new __holder(null);   // Current function occurrence.
  var oobjectdummy = null;   // Object for intermediate storing.
  var bchange = false;   // Indicator flag in case of change.
  var bexpr1 = false;   // Ausdruck1
  var bexpr2 = false;   // Ausdruck2

  var ofuncoccs = ocurrentmodel.ObjOccListFilter(Constants.OT_FUNC);   // List of function occurrences.
  var ofuncoccsarray = __redimArray(ofuncoccsarray, new __ArrayInfo(0, ofuncoccs.length), null, false);  __createArray(null);   // List of function occurrences for exchange.
  if (ofuncoccs.length > 0) {
    for (var i = 0 ; i < (ofuncoccs.length - 1)+1 ; i++ ){
      ofuncoccsarray[__getArrayIndex(ofuncoccsarray,1,i)] = ofuncoccs[i];
    }

    // Sorting on the basis of the size.
    while (bchange == false) {
      bchange = __toBool(true);
      for (var j = 0 ; j < (ofuncoccs.length - 2)+1 ; j++ ){
        bexpr1 = __toBool(vbIsNumeric(ofuncoccsarray[__getArrayIndex(ofuncoccsarray,1,j)].ObjDef().Attribute(Constants.AT_USER_ATTR5, g_nloc).GetValue(true)));
        bexpr2 = __toBool(vbIsNumeric(ofuncoccsarray[__getArrayIndex(ofuncoccsarray,1,(j + 1))].ObjDef().Attribute(Constants.AT_USER_ATTR5, g_nloc).GetValue(true)));
        if (bexpr1 && bexpr2) {
          if (__toInteger(ofuncoccsarray[__getArrayIndex(ofuncoccsarray,1,j)].ObjDef().Attribute(Constants.AT_USER_ATTR5, g_nloc).GetValue(true)) > __toInteger(ofuncoccsarray[__getArrayIndex(ofuncoccsarray,1,(j + 1))].ObjDef().Attribute(Constants.AT_USER_ATTR5, g_nloc).GetValue(true))) {
            oobjectdummy = ofuncoccsarray[__getArrayIndex(ofuncoccsarray,1,j)];
            ofuncoccsarray[__getArrayIndex(ofuncoccsarray,1,j)] = ofuncoccsarray[__getArrayIndex(ofuncoccsarray,1,(j + 1))];
            ofuncoccsarray[__getArrayIndex(ofuncoccsarray,1,(j + 1))] = oobjectdummy;
            oobjectdummy = null;
            bchange = __toBool(false);
          }
        } else if (bexpr1 || bexpr2) {
          if (! (bexpr1) && bexpr2) {
            oobjectdummy = ofuncoccsarray[__getArrayIndex(ofuncoccsarray,1,j)];
            ofuncoccsarray[__getArrayIndex(ofuncoccsarray,1,j)] = ofuncoccsarray[__getArrayIndex(ofuncoccsarray,1,(j + 1))];
            ofuncoccsarray[__getArrayIndex(ofuncoccsarray,1,(j + 1))] = oobjectdummy;
            oobjectdummy = null;
            bchange = __toBool(false);
          }
        }
      }
    }

    // Output of sorted functions.
    for (var i = 0 ; i < (ofuncoccs.length - 1)+1 ; i++ ){
      ocurrentfuncocc.value = ofuncoccsarray[__getArrayIndex(ofuncoccsarray,1,i)];
      // Output will only be executed if the attribute is maintained.
      if (ocurrentfuncocc.value.SymbolNum() != Constants.ST_PRCS_IF && (vbIsNumeric(ocurrentfuncocc.value.ObjDef().Attribute(Constants.AT_USER_ATTR5, g_nloc).GetValue(true)) == true))
      {
        outfuncdata(ocurrentfuncocc, ndepth, ncurrentdepth, nmodellevel, new __holder(__toString((i + 1))), oasprozmodellist, ssourcefuncproc, new __holder(""));
      }
      ocurrentfuncocc.value = null;
    }
  }
}


// --------------------------------------------------------------
// Subprogram OutObjOfModels for outputting certain object types that have relationships to the functions in the current model.

function outobjofmodels(oobjdefs, sheadline)
{
  if (oobjdefs.length > 0) {
    oobjdefs = __sort(oobjdefs, Constants.AT_NAME, Constants.SORT_NONE, Constants.SORT_NONE, g_nloc);

    g_ooutfile.value.OutputLn("", getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0);
    g_ooutfile.value.OutputLn(sheadline, getString("TEXT48"), 14, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0);
    g_ooutfile.value.OutputLn("", getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0);
    for (var i = 0 ; i < (oobjdefs.length - 1)+1 ; i++ ){
      g_ooutfile.value.OutputLn(oobjdefs[i].Name(g_nloc), getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
      var oattributes = oobjdefs[i].AttrList(g_nloc);
      if (oattributes.length > 0) {
        g_ooutfile.value.BeginTable(100, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
        outofattributes(oattributes);
        g_ooutfile.value.EndTable("", 100, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
        g_ooutfile.value.OutputLn("", getString("TEXT48"), 8, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0);
      }
    }
  }
}


// --------------------------------------------------------------
// OutOfAssignedModels for outputting the assigned models.
// oCurrentFuncObjOcc = Current function occurrence.
// oAsProzModelList = List of assigned models (of the Processes type).
// nCurrentDepth = Current depth of the model levels.
// nModelLevel = Model number on the current level.

function outofassignedmodels(ocurrentfuncobjocc, oasprozmodellist, ndepth, ncurrentdepth, nmodellevel)
{
  var oassignedmodels = null;   // List of models which are assigned to the function.
  var osubmodfunc = null;   // Object definition of the current function definition in the assigned model.
  var soutstring = new __holder(""); 
  var bcheck = false; 

  oassignedmodels = ocurrentfuncobjocc.value.ObjDef().AssignedModels();
  oassignedmodels = __sort(oassignedmodels, Constants.AT_NAME, Constants.SORT_NONE, Constants.SORT_NONE, g_nloc);

  if (oassignedmodels.length > 0) {
    var bColored = true;   // variable to change background color of table rows
      
    g_ooutfile.value.TableRow();
    g_ooutfile.value.TableCell(txtAssignedModels + ": ", 50, getString("TEXT48"), 12, Constants.C_BLACK, getTableCellColor_Bk(false), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.value.TableCell("", 50, getString("TEXT48"), 12, Constants.C_BLACK, getTableCellColor_Bk(false), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    for (var h = 0 ; h < (oassignedmodels.length - 1)+1 ; h++ ){
      bcheck = __toBool(checkassignedmodel(oassignedmodels[h], oasprozmodellist, ndepth, ncurrentdepth, nmodellevel, soutstring));
      g_ooutfile.value.TableRow();
      g_ooutfile.value.TableCell(soutstring.value, 50, getString("TEXT48"), 12, Constants.C_BLACK, getTableCellColor_Bk(bColored), 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
      g_ooutfile.value.TableCell(oassignedmodels[h].Type(), 50, getString("TEXT48"), 12, Constants.C_BLACK, getTableCellColor_Bk(bColored), 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    
      bColored = !bColored; // Change background color
    }
  }
}


// --------------------------------------------------------------
// OutOfAttributes for outputting the allowed attributes.
// oAttributes = List of maintained attributes.

function outofattributes(oattributes)
{
  var bAttrColored = true;   // variable to change background color of table rows (attributes)                            

  g_ooutfile.value.TableRow();
  g_ooutfile.value.TableCell(getString("TEXT53"), 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
  g_ooutfile.value.TableCell("", 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
  for (var i = 0 ; i < (oattributes.length - 1)+1 ; i++ ){
    var ocurrentattribute = oattributes[i];
    g_ooutfile.value.TableRow();
    g_ooutfile.value.TableCell(ocurrentattribute.Type(), 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);

    if (ocurrentattribute.TypeNum() == Constants.AT_DESC) {
      // Line breaks NOT removed
      g_ooutfile.value.TableCell(ocurrentattribute.GetValue(false), 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    } else {
      g_ooutfile.value.TableCell(ocurrentattribute.GetValue(true), 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    }

    bAttrColored = !bAttrColored; // Change background color (attributes)    
  }
}

function _outputFunctionAttributes(oFunctionDef)
{             

    var blnIsControlActivity = oFunctionDef.Attribute(Constants.AT_USER_ATTR_BOOL_100,g_nloc).GetValue(false);
    //var blnIsControlActivity = oFunctionDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("317e1a3f-7df8-11da-0c60-cf8e338f9a0b"),g_nloc).GetValue(false);

    
    
    if(blnIsControlActivity == "True" || blnIsControlActivity == "True (EN)") //custom 202202141440
        _writeControlActivity(oFunctionDef);
    else
        _writeNormalActivity(oFunctionDef);
}

function _writeControlActivity(oFunctionDef)
{  
    g_ooutfile.value.TableRow();
    g_ooutfile.value.TableCell("Control activity attributes:", 100, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_RED, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    //g_ooutfile.value.TableCell("", 0, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_RED, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
  
    g_ooutfile.value.TableRow();
    g_ooutfile.value.TableCell("Order No.", 22, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.value.TableCell(oFunctionDef.Attribute(Constants.AT_USER_ATTR5,g_nloc).GetValue(true), 78, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
  
    //g_ooutfile.value.TableRow();
    //g_ooutfile.value.TableCell("Type", 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    //g_ooutfile.value.TableCell(oFunctionDef.Attribute(Constants.AT_TYPE_6,g_nloc).GetValue(true), 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    
    //g_ooutfile.value.TableRow();
    //g_ooutfile.value.TableCell("Control activity", 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    //g_ooutfile.value.TableCell("True", 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    
    g_ooutfile.value.TableRow();
    g_ooutfile.value.TableCell("Name", 22, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.value.TableCell(oFunctionDef.Attribute(Constants.AT_NAME,g_nloc).GetValue(true), 78, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    
    g_ooutfile.value.TableRow();
    g_ooutfile.value.TableCell("Long Description", 22, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.value.TableCell(oFunctionDef.Attribute(Constants.AT_DESC,g_nloc).GetValue(true), 78, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    
    g_ooutfile.value.TableRow();
    g_ooutfile.value.TableCell("Control frequency", 22, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.value.TableCell(oFunctionDef.Attribute(Constants.AT_AAM_CTRL_FREQUENCY,g_nloc).GetValue(true), 78, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    
    g_ooutfile.value.TableRow();
    g_ooutfile.value.TableCell("Control type", 22, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.value.TableCell(oFunctionDef.Attribute(Constants.AT_AAM_CTRL_EFFECT,g_nloc).GetValue(true), 78, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    
    g_ooutfile.value.TableRow();
    g_ooutfile.value.TableCell("Control evidence", 22, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.value.TableCell(oFunctionDef.Attribute(Constants.AT_UA_TXT_100,g_nloc).GetValue(true), 78, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    
    g_ooutfile.value.TableRow();
    g_ooutfile.value.TableCell("Type (control)", 22, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.value.TableCell(oFunctionDef.Attribute(Constants.AT_UA_VAL_221,g_nloc).GetValue(true), 78, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    
    g_ooutfile.value.TableRow();
    g_ooutfile.value.TableCell("Control ID", 22, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.value.TableCell(oFunctionDef.Attribute(Constants.AT_AAM_CTRL_ID,g_nloc).GetValue(true), 78, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    
    g_ooutfile.value.TableRow();
    g_ooutfile.value.TableCell("Key control", 22, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.value.TableCell(oFunctionDef.Attribute(Constants.AT_AAM_KEY_CTRL,g_nloc).GetValue(true), 78, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    
}

function _writeNormalActivity(oFunctionDef)
{   
    g_ooutfile.value.TableRow();
    g_ooutfile.value.TableCell("Activity attributes:", 100, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    //g_ooutfile.value.TableCell("", 0, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
  
    g_ooutfile.value.TableRow();
    g_ooutfile.value.TableCell("Order No.", 22, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.value.TableCell(oFunctionDef.Attribute(Constants.AT_USER_ATTR5,g_nloc).GetValue(true), 78, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
  
    //g_ooutfile.value.TableRow();
    //g_ooutfile.value.TableCell("Type", 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    //g_ooutfile.value.TableCell(oFunctionDef.Attribute(Constants.AT_TYPE_6,g_nloc).GetValue(true), 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    
    g_ooutfile.value.TableRow();
    g_ooutfile.value.TableCell("Name", 22, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.value.TableCell(oFunctionDef.Attribute(Constants.AT_NAME,g_nloc).GetValue(true), 78, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    
    g_ooutfile.value.TableRow();
    g_ooutfile.value.TableCell("Long Description", 22, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.value.TableCell(oFunctionDef.Attribute(Constants.AT_DESC,g_nloc).GetValue(true), 78, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    
}

function _writePrecedingEvents(oFunctionDef)
{             
    var objPrecedingEvent = getActivatesEvent(oFunctionDef);
    var txtEventStatus = "Status";
    
    if(g_blnIsStartEvent)
        txtEventStatus = "Start";
    else
        txtEventStatus = "Status";
        
    
    if(objPrecedingEvent != null)
        writeEventInfo(objPrecedingEvent, txtEventStatus);
    else
    {
        var objPrecOperator = getActivatesOperator(oFunctionDef);
        var objEvents = new Array();
        
        if(objPrecOperator != null)
        {
            getEvaluatedByEvents(objPrecOperator, oFunctionDef, objEvents);
            
            var objLinkedOperator = getLinksOperator(objPrecOperator);
                
            if(objLinkedOperator != null)
                getEvaluatedByEvents(objLinkedOperator, oFunctionDef, objEvents);
                
            if(objEvents.length > 0)
            {
                if(objLinkedOperator != null)
                    writeEventsInfo(objEvents, txtEventStatus,objLinkedOperator);
                else
                    writeEventsInfo(objEvents, txtEventStatus,objPrecOperator);
            }
        }
    }
    
    g_blnIsStartEvent = false;
}

function getActivatesEvent(p_Object)
{
	var objTypes = new Array();
	objTypes[0] = Constants.OT_EVT;
	var conTypes = new Array();
	conTypes[0] = Constants.CT_ACTIV_1;
    conTypes[1] = Constants.CT_ACTIV_2;
    conTypes[2] = Constants.CT_ACTIV_3;
    //conTypes[3] = Constants.CT_ACTIVATES; //Custom 20220201
	
	var objConnectedObjects = p_Object.getConnectedObjs(objTypes, Constants.EDGES_IN, conTypes);
    
    if(objConnectedObjects.length > 0)
    {
        var objEvent = objConnectedObjects[0];
        
        if(isEventPositionAbove(objEvent,p_Object))
            return objEvent;
        else
            return null;
    }
    else
        return null;
}

function getActivatesOperator(p_Object)
{
	var objTypes = new Array();
	objTypes[0] = Constants.OT_RULE;
	var conTypes = new Array();
	conTypes[0] = Constants.CT_ACTIV_1;
    conTypes[1] = Constants.CT_ACTIV_2;
    conTypes[2] = Constants.CT_ACTIV_3;
    //conTypes[3] = Constants.CT_ACTIVATES; //Custom 20220201
	
	var objConnectedObjects = p_Object.getConnectedObjs(objTypes, Constants.EDGES_IN, conTypes);
    
    if(objConnectedObjects.length > 0)
        return objConnectedObjects[0];
    else
        return null;
}

function getLinksOperator(p_Object)
{
	var objTypes = new Array();
	objTypes[0] = Constants.OT_RULE;
	var conTypes = new Array();
	conTypes[0] = Constants.CT_LNK_1;
    conTypes[1] = Constants.CT_LNK_2;
    conTypes[2] = Constants.CT_LNK_3;
    conTypes[3] = Constants.CT_LNK_4;
	
	var objConnectedObjects = p_Object.getConnectedObjs(objTypes, Constants.EDGES_IN, conTypes);
    
    if(objConnectedObjects.length > 0)
        return objConnectedObjects[0];
    else
        return null;
}

function getEvaluatedByEvents(p_Object, objFunctionDef, objAboveEventsList)
{
	var objTypes = new Array();
	objTypes[0] = Constants.OT_EVT;
	var conTypes = new Array();
	conTypes[0] = Constants.CT_IS_EVAL_BY_1;
	
	var objConnectedObjects = p_Object.getConnectedObjs(objTypes, Constants.EDGES_IN, conTypes);
    
    if(objConnectedObjects.length > 0)
    {   
        for(var index=0; index<objConnectedObjects.length; index++) 
        {
            if(isEventPositionAbove(objConnectedObjects[index],objFunctionDef))
                objAboveEventsList[objAboveEventsList.length] = objConnectedObjects[index];
        }
    }
}

function isEventPositionAbove(objEvent, objObject)
{
    var objModels = new Array();
	objModels[0] = g_objCurrentModel;
    
    objEventOccList = objEvent.OccList(objModels);
    objObjectOcc = objObject.OccList(objModels)[0];
    
    for(var i=0; i<objEventOccList.length; i++) 
    {
        if(objEventOccList[i].Y() < objObjectOcc.Y())
            return true;
    }
    
    return false;
}

function writeEventInfo(objEvent, txtEventStatus)
{
    g_ooutfile.value.OutputLnF("", "REPORT3");
    g_ooutfile.value.BeginTable(100, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
    
    g_ooutfile.value.TableRow();
    g_ooutfile.value.TableCell(txtEventStatus + " event attributes:", 100, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    //g_ooutfile.value.TableCell("", 0, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
  
    //g_ooutfile.value.TableRow();
    //g_ooutfile.value.TableCell("Type", 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    //g_ooutfile.value.TableCell("Event (Status)", 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
  
    g_ooutfile.value.TableRow();
    g_ooutfile.value.TableCell("Name", 22, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.value.TableCell(objEvent.Attribute(Constants.AT_NAME,g_nloc).GetValue(true), 78, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    
    g_ooutfile.value.EndTable("", 100, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
}

function writeEventsInfo(objEvents, txtEventStatus, objPrecOperator)
{
    if(objEvents.length == 1)
        writeEventInfo(objEvents[0], txtEventStatus);
    else
    {
        var objModels = new Array();
	    objModels[0] = g_objCurrentModel;
    
        var objPrecOperatorOcc = objPrecOperator.OccList(objModels)[0];
        
        var txtOperator = " ";
        if(objPrecOperatorOcc.SymbolNum() == Constants.ST_OPR_AND_1 || objPrecOperatorOcc.SymbolNum() == Constants.ST_OPR_XOR_1)
            txtOperator = " and ";
        else
        {
            if(objPrecOperatorOcc.SymbolNum() == Constants.ST_OPR_OR_1)
                txtOperator = " or ";
        }
        
        g_ooutfile.value.OutputLnF("", "REPORT3");
        g_ooutfile.value.BeginTable(100, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
      
        g_ooutfile.value.TableRow();
        g_ooutfile.value.TableCell(txtEventStatus + " event attributes:", 100, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
        //g_ooutfile.value.TableCell("", 0, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
  
        //g_ooutfile.value.TableRow();
        //g_ooutfile.value.TableCell("Type", 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
        //g_ooutfile.value.TableCell("Event (Status)", 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
        
        var text = objEvents[0].Attribute(Constants.AT_NAME,g_nloc).GetValue(true);
        for(var index=1; index<objEvents.length; index++) 
        {
            text = text + txtOperator;
            text = text + objEvents[index].Attribute(Constants.AT_NAME,g_nloc).GetValue(true);
        }
        
        g_ooutfile.value.TableRow();
        g_ooutfile.value.TableCell("Name", 22, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
        g_ooutfile.value.TableCell(text, 78, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
        
        g_ooutfile.value.EndTable("", 100, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    }
            
}

function writeEndEvents(objModel)
{
    var objEventsList = objModel.ObjDefListFilter(Constants.OT_EVT);
    
    var objEndEventList = getEndEvents(objEventsList, objModel);
    
    if(objEndEventList.length == 0)
        return;
    
    if(objEndEventList.length == 1)
        writeEventInfo(objEndEventList[0], "End");
    else
    {
        g_ooutfile.value.OutputLnF("", "REPORT3");
        g_ooutfile.value.BeginTable(100, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
      
        g_ooutfile.value.TableRow();
        g_ooutfile.value.TableCell("End event attributes:", 100, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
        //g_ooutfile.value.TableCell("", 0, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
  
        //g_ooutfile.value.TableRow();
        //g_ooutfile.value.TableCell("Type", 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
        //g_ooutfile.value.TableCell("Event (End)", 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
        
        var text = objEndEventList[0].Attribute(Constants.AT_NAME,g_nloc).GetValue(true);
        var txtOperator = "";
    
        for(var i=1; i<objEndEventList.length; i++) 
        {
            txtOperator = getEndEventTxtOperator(objEndEventList[i], objModel);
            text = text + txtOperator;
            text = text + objEndEventList[i].Attribute(Constants.AT_NAME,g_nloc).GetValue(true);
        }
        
        g_ooutfile.value.TableRow();
        g_ooutfile.value.TableCell("Name", 22, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
        g_ooutfile.value.TableCell(text, 78, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
        
        g_ooutfile.value.EndTable("", 100, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    } 
    
}

function getEndEvents(objEventList, objModel)
{
    var objEventOccList;
    var objEventOcc;
    var objEndEventsList = new Array();
    var objModels = new Array();
    objModels[0] = g_objCurrentModel;
    
    for(var i=0; i<objEventList.length; i++) 
    {
        objEventOccList = objEventList[i].OccList(objModels);
        
        for(var j=0; j<objEventOccList.length; j++)
        {
            objEventOcc = objEventOccList[j];
            
            if(objectHasNoOutputConnections(objEventOcc))
                objEndEventsList[objEndEventsList.length] = objEventList[i];
            else
            {
                if(objectIsConnectedToEndProcessInterface(objEventOcc))
                    objEndEventsList[objEndEventsList.length] = objEventList[i];
                else
                {
                    if(objectIsConnectedToEndOperator(objEventOcc))
                        objEndEventsList[objEndEventsList.length] = objEventList[i];
                }
            }
        }
    }
    
    return objEndEventsList;
}

function objectHasNoOutputConnections(objOcc)
{
    return (objOcc.OutEdges(Constants.EDGES_ALL).length == 0);
}

function objectIsConnectedToEndProcessInterface(objOcc)
{
    var SymbolTypes = new Array();
    SymbolTypes[0] = Constants.ST_PRCS_IF;
            
    var objProcIfOccList = objOcc.getConnectedObjOccs(SymbolTypes, Constants.EDGES_OUT);
            
    if(objProcIfOccList.length != 0)
    {
        var objProcIfOcc = objProcIfOccList[0];
                
        if(objProcIfOcc.OutEdges(Constants.EDGES_ALL).length == 0)
            return true;
    }
    
    return false;
}

function objectIsConnectedToEndOperator(objOcc)
{
    var SymbolTypes = new Array();
    SymbolTypes[0] = Constants.ST_OPR_AND_1;
    SymbolTypes[1] = Constants.ST_OPR_OR_1;
    SymbolTypes[2] = Constants.ST_OPR_XOR_1;
            
    var objOperatorOccList = objOcc.getConnectedObjOccs(SymbolTypes, Constants.EDGES_OUT);
                
    if(objOperatorOccList.length != 0)
    {
        var objOperatorOcc = getLastOperator(objOperatorOccList[0]);
        
        if(objectIsConnectedToEndProcessInterface(objOperatorOcc))
            return true;
    }
    
    return false;
}

function getLastOperator(objOcc)
{
    var SymbolTypes = new Array();
    SymbolTypes[0] = Constants.ST_OPR_AND_1;
    SymbolTypes[1] = Constants.ST_OPR_OR_1;
    SymbolTypes[2] = Constants.ST_OPR_XOR_1;
            
    var objOperatorOccList = objOcc.getConnectedObjOccs(SymbolTypes, Constants.EDGES_OUT);
    
    if(objOperatorOccList.length != 0)
        return getLastOperator(objOperatorOccList[0]);
    else
        return objOcc;
}

function getEventStatus(objEvent)
{
    var database = ArisData.getActiveDatabase();
    var filter = database.ActiveFilter();
    var checkAttribute = true;
    var attributeEventStatus;
    
    try
    {
      attributeEventStatus = filter.UserDefinedAttributeTypeNum("4c75c4a0-a694-11dd-7812-0017a4cd8aa8");
    }
     catch(e)
    {
      checkAttribute = false;
    }
    
    if(checkAttribute == true)
    {
        if(objEvent.Attribute(attributeEventStatus, g_nloc).getValue() != "")
            return objEvent.Attribute(attributeEventStatus, g_nloc).getValue();
        else
            return "Status";
    }
    else
        return "Status";
}

function getEndEventTxtOperator(objEndEvent, objModel)
{
    var objModels = new Array();
	objModels[0] = objModel;
    var txtOperator = " and ";
    var objEventOcc = objEndEvent.OccList(objModels)[0];
    var connectedOperatorBellowOcc = getOperatorBellow(objEventOcc);
    
    if(connectedOperatorBellowOcc != null)
    {
        if(connectedOperatorBellowOcc.SymbolNum() == Constants.ST_OPR_AND_1)
            txtOperator = " and ";
        else
        {
            if(connectedOperatorBellowOcc.SymbolNum() == Constants.ST_OPR_OR_1 || connectedOperatorBellowOcc.SymbolNum() == Constants.ST_OPR_XOR_1)
                txtOperator = " or ";
        }
    }
    
    return txtOperator;
}

function getOperatorBellow(objEventOcc)
{
	var SymbolTypes = new Array();
    SymbolTypes[0] = Constants.ST_OPR_AND_1;
    SymbolTypes[1] = Constants.ST_OPR_OR_1;
    SymbolTypes[2] = Constants.ST_OPR_XOR_1;
	
	var objConnectedObjectsOcc = objEventOcc.getConnectedObjOccs(SymbolTypes, Constants.EDGES_OUT);
    
    if(objConnectedObjectsOcc.length > 0)
        return objConnectedObjectsOcc[0];
    else
        return null;
}

function _outputRelationships(oFunctionDef)
{
    var blnIsControlActivity = oFunctionDef.Attribute(Constants.AT_USER_ATTR_BOOL_100,g_nloc).GetValue(false);  
    var blnTextColor;
    
    if(blnIsControlActivity == "True" || blnIsControlActivity == "True (EN)") //custom 202202141440
        blnTextColor = Constants.C_BLACK;
    else
        blnTextColor = Constants.C_BLACK;
    
    g_ooutfile.value.TableRow();
    g_ooutfile.value.TableCell(txtRelations + ": ", 100, getString("TEXT48"), 12, blnTextColor, getTableCellColor_Bk(false), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    //g_ooutfile.value.TableCell("", 0, getString("TEXT48"), 12, blnTextColor, getTableCellColor_Bk(false), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    
    writeApplications(oFunctionDef, blnTextColor);  
    writeOrgUnitsPositions(oFunctionDef, blnTextColor);   
    writeInputDocs(oFunctionDef, blnTextColor);
    writeOutputDocs(oFunctionDef, blnTextColor);
    writeRisks(oFunctionDef, blnTextColor);
    writeDetails(oFunctionDef, blnTextColor);     
}

function writeApplications(p_Object, blnTextColor)
{
	var outText = "";
	var objTypes = new Array();
	objTypes[0] = Constants.OT_APPL_SYS_TYPE;
	objTypes[1] = Constants.OT_INFO_CARR;
	objTypes[2] = Constants.OT_LST;
	objTypes[3] = Constants.OT_LST_DSGN;
	var conTypes = new Array();
	conTypes[0] = Constants.CT_SUPP_1;
	conTypes[1] = Constants.CT_SUPP_2;
	conTypes[2] = Constants.CT_SUPP_3;
	conTypes[3] = Constants.CT_SUPPORTS;
	conTypes[4] = Constants.CT_CAN_SUPP_1;
	conTypes[5] = Constants.CT_CAN_SUPP_2;
	conTypes[6] = Constants.CT_PROV_INP_FOR;
	
	var objConnectedApplicationsList = p_Object.getConnectedObjs(objTypes, Constants.EDGES_INOUTASSIGN,conTypes);
    
    if(objConnectedApplicationsList.length > 0)
    {
        g_ooutfile.value.TableRow();
        g_ooutfile.value.TableCell("Application/Media", 22, getString("TEXT48"), 12, blnTextColor, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0); //custom 202202141440
    
	    for(var index=0; index<objConnectedApplicationsList.length; index++) 
        {
            outText = objConnectedApplicationsList[index].Attribute(Constants.AT_NAME, g_nloc).GetValue(true);
            
            if(index==0)
            {
                g_ooutfile.value.TableCell(outText, 78, getString("TEXT48"), 12, blnTextColor, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
            }
            else
            {
                g_ooutfile.value.TableRow();
                g_ooutfile.value.TableCell("", 22, getString("TEXT48"), 12, blnTextColor, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
	            g_ooutfile.value.TableCell(outText, 78, getString("TEXT48"), 12, blnTextColor, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
            }
        }
    }
}

function writeOrgUnitsPositions(p_Object, blnTextColor) 
{
	var outText = "";
	var objTypes = new Array();
	objTypes[0] = Constants.OT_ORG_UNIT;
	objTypes[1] = Constants.OT_POS;
	objTypes[2] = Constants.OT_PERS;
	objTypes[3] = Constants.OT_PERS_TYPE;
	objTypes[4] = Constants.OT_ORG_UNIT_TYPE;
	var conTypes = new Array();
	conTypes[0] = Constants.CT_CONTR_TO_1;
	conTypes[1] = Constants.CT_CONTR_TO_2;
	conTypes[2] = Constants.CT_EXEC_1;
	conTypes[3] = Constants.CT_EXEC_2;
	//conTypes[4] = Constants.CT_EXEC_3; //Custom 20220201
	conTypes[4] = Constants.CT_EXEC_4;
	conTypes[5] = Constants.CT_EXEC_5;
	
	var objConnectedOrg = p_Object.getConnectedObjs(objTypes, Constants.EDGES_INOUTASSIGN, conTypes);
    
    if(objConnectedOrg.length > 0)
    {
        g_ooutfile.value.TableRow();
        g_ooutfile.value.TableCell("Organizational unit/Position", 22, getString("TEXT48"), 12, blnTextColor, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    
	    for(var index=0; index<objConnectedOrg.length; index++) 
        {
            outText = objConnectedOrg[index].Attribute(Constants.AT_NAME, g_nloc).GetValue(true);
            
            if(index==0)
            {
                g_ooutfile.value.TableCell(outText, 78, getString("TEXT48"), 12, blnTextColor, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
            }
            else
            {
                g_ooutfile.value.TableRow();
                g_ooutfile.value.TableCell("", 22, getString("TEXT48"), 12, blnTextColor, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
	            g_ooutfile.value.TableCell(outText, 78, getString("TEXT48"), 12, blnTextColor, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
            }
        }
    }
    
}

function writeInputDocs(p_Object, blnTextColor) 
{
	var outText = "";
	var objTypes = new Array();
	objTypes[0] = Constants.OT_TECH_TRM;
	objTypes[1] = Constants.OT_INFO_CARR;
	objTypes[2] = Constants.OT_LST;
	objTypes[3] = Constants.OT_LST_DSGN;
	var conTypes = new Array();
	conTypes[0] = Constants.CT_IS_INP_FOR;
	
	var objConnectedInputDocsList = p_Object.getConnectedObjs(objTypes, Constants.EDGES_IN, conTypes);
    
	if(objConnectedInputDocsList.length > 0)
    {
        g_ooutfile.value.TableRow();
        g_ooutfile.value.TableCell("Input documents", 22, getString("TEXT48"), 12, blnTextColor, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    
	    for(var index=0; index<objConnectedInputDocsList.length; index++) 
        {
            outText = objConnectedInputDocsList[index].Attribute(Constants.AT_NAME, g_nloc).GetValue(true);
            
            if(index==0)
            {
                g_ooutfile.value.TableCell(outText, 78, getString("TEXT48"), 12, blnTextColor, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
            }
            else
            {
                g_ooutfile.value.TableRow();
                g_ooutfile.value.TableCell("", 22, getString("TEXT48"), 12, blnTextColor, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
	            g_ooutfile.value.TableCell(outText, 78, getString("TEXT48"), 12, blnTextColor, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
            }
        }
    }
}

function writeOutputDocs(p_Object, blnTextColor) 
{
	var outText = "";
	var objTypes = new Array();
	objTypes[0] = Constants.OT_TECH_TRM;
	objTypes[1] = Constants.OT_INFO_CARR;
	objTypes[2] = Constants.OT_LST;
	objTypes[3] = Constants.OT_LST_DSGN;
	var conTypes = new Array();
	conTypes[0] = Constants.CT_HAS_OUT;
	
	var objConnectedOutputDocsList = p_Object.getConnectedObjs(objTypes, Constants.EDGES_OUT, conTypes);
	
    if(objConnectedOutputDocsList.length > 0)
    {
        g_ooutfile.value.TableRow();
        g_ooutfile.value.TableCell("Output documents", 22, getString("TEXT48"), 12, blnTextColor, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    
	    for(var index=0; index<objConnectedOutputDocsList.length; index++) 
        {
            outText = objConnectedOutputDocsList[index].Attribute(Constants.AT_NAME, g_nloc).GetValue(true);
            
            if(index==0)
            {
                g_ooutfile.value.TableCell(outText, 78, getString("TEXT48"), 12, blnTextColor, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
            }
            else
            {
                g_ooutfile.value.TableRow();
                g_ooutfile.value.TableCell("", 22, getString("TEXT48"), 12, blnTextColor, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
	            g_ooutfile.value.TableCell(outText, 78, getString("TEXT48"), 12, blnTextColor, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
            }
        }
    }
}

function writeRisks(p_Object, blnTextColor) 
{
	var outText = "";
	var objTypes = new Array();
	objTypes[0] = Constants.OT_RISK;
	var conTypes = new Array();
	conTypes[0] = Constants.CT_OCCUR;
    conTypes[1] = Constants.CT_OCCURS_AT;
	
	var objConnectedObj = p_Object.getConnectedObjs(objTypes, Constants.EDGES_INOUTASSIGN, conTypes);
    
    if(objConnectedObj.length > 0)
    {
        g_ooutfile.value.TableRow();
        g_ooutfile.value.TableCell("Risks", 22, getString("TEXT48"), 12, blnTextColor, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    
	    for(var index=0; index<objConnectedObj.length; index++) 
        {
            outText = objConnectedObj[index].Attribute(Constants.AT_NAME, g_nloc).GetValue(true);
            
            if(index==0)
            {
                g_ooutfile.value.TableCell(outText, 78, getString("TEXT48"), 12, blnTextColor, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
            }
            else
            {
                g_ooutfile.value.TableRow();
                g_ooutfile.value.TableCell("", 22, getString("TEXT48"), 12, blnTextColor, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
	            g_ooutfile.value.TableCell(outText, 78, getString("TEXT48"), 12, blnTextColor, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
            }
        }
    }
    
}

function writeDetails(p_Object, blnTextColor) 
{
	var outText = "";
	var objTypes = new Array();
	objTypes[0] = Constants.OT_LST_DSGN;
	var conTypes = new Array();
	conTypes[0] = Constants.CT_CAN_USE_1;
	conTypes[1] = Constants.CT_CAN_USE_2;
    conTypes[2] = Constants.CT_CAN_CRT;
	
	var objConnectedDetailsList = p_Object.getConnectedObjs(objTypes, Constants.EDGES_OUT, conTypes);
	
    if(objConnectedDetailsList.length > 0)
    {
        g_ooutfile.value.TableRow();
        g_ooutfile.value.TableCell("Comment", 22, getString("TEXT48"), 12, blnTextColor, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    
	    for(var index=0; index<objConnectedDetailsList.length; index++) 
        {
            outText = objConnectedDetailsList[index].Attribute(Constants.AT_NAME, g_nloc).GetValue(true);
            
            if(index==0)
            {
                g_ooutfile.value.TableCell(outText, 78, getString("TEXT48"), 12, blnTextColor, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
            }
            else
            {
                g_ooutfile.value.TableRow();
                g_ooutfile.value.TableCell("", 22, getString("TEXT48"), 12, blnTextColor, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
	            g_ooutfile.value.TableCell(outText, 78, getString("TEXT48"), 12, blnTextColor, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
            }
        }
    }
}

function outofepkfunc(ocurrentmodel, ocurrentoccfunc, ndepth, ncurrentdepth, nmodellevel, oasprozmodellist, ssourcefuncproc, spredfuncocc, bisindex)
{
  if(g_nOptOutput == 0) {
    outofepkfunc_std(ocurrentmodel, ocurrentoccfunc, ndepth, ncurrentdepth, nmodellevel, oasprozmodellist, ssourcefuncproc, spredfuncocc, bisindex);
  } else if(g_nOptOutput == 1) {
    outofepkfunc_iso(ocurrentmodel, ocurrentoccfunc, ndepth, ncurrentdepth, nmodellevel, oasprozmodellist, ssourcefuncproc, spredfuncocc, bisindex);
  }
}

function outofepkfunc_iso(ocurrentmodel, ocurrentoccfunc, ndepth, ncurrentdepth, nmodellevel, oasprozmodellist, ssourcefuncproc, spredfuncocc, bisindex)
{
  var onextobjocclist = __createArray(null); 
  var currentcxnocc = null; 
  var otarfulist = new __holder(null); 
  var nobjabspos = __createArray(0.0); 
  var bdonefunc = false; 
  var sidofevaluatefun = new __holder(""); 

  var srulestring = new __holder(""); 
  var sdummy = new __holder(""); 
  var ndummy = new __holder(0); 
  var x = 0.0;;   var y = 0.0;;   var z = 0.0; 

  // Default settings
  var lnumberoftargetfuncobjects = undefined; 
  lnumberoftargetfuncobjects = 0;
  onextobjocclist = __redimArray(onextobjocclist, new __ArrayInfo(0, 0), null, false); 
  otarfulist.value = new Array();

  switch(spredfuncocc) {
    case "":
      srulestring.value = __toString("");
    break;
    default:
      if (vbIsNumeric(spredfuncocc) && bisindex) {
        srulestring.value = __toString(g_ssourceruleoffunc[__getArrayIndex(g_ssourceruleoffunc,1,__toInteger(spredfuncocc))]);
      }
      else {
        srulestring.value = __toString(((((((((" "+txtFunction+" " + spredfuncocc) + " ("+txtChapter+" ") + vbStr(ncurrentdepth.value)) + ".") + vbStr(nmodellevel.value)) + ".") + vbStr(g_nprozfunclevel)) + ")"));
      }
    }
  g_nprozfunclevel = __toInteger((g_nprozfunclevel + 1));
  findnextfunc(ocurrentmodel, ocurrentoccfunc.value, otarfulist);
  bdonefunc = __toBool(checkobj(ocurrentoccfunc, sdummy, ndummy, ndummy, sdummy, false, false, g_odonemodfuncoccs, g_sdfuoccid, ocurrentoccfunc.value.ObjDef().Name(g_nloc)));
  if (otarfulist.value.length > 0) {
    nobjabspos = __redimArray(nobjabspos, new __ArrayInfo(0, (otarfulist.value.length - 1)), 0.0, false); 
    onextobjocclist = __redimArray(onextobjocclist, new __ArrayInfo(0, (otarfulist.value.length - 1)), null, false); 
    for (var j = 0 ; j < (otarfulist.value.length - 1)+1 ; j++ ){
      onextobjocclist[__getArrayIndex(onextobjocclist,1,j)] = otarfulist.value[j];
      x = otarfulist.value[j].X();
      y = otarfulist.value[j].Y();
      z = __toDouble((((x * x)) + ((y * y))));
      nobjabspos[__getArrayIndex(nobjabspos,1,j)] = Math.sqrt(z);
    }

    // The paths are inserted in the global list to make sure that they are taken into account when sorting.
    g_ssortstrings = __redimArray(g_ssortstrings, new __ArrayInfo(0, (otarfulist.value.length - 1)), "", true); 
    for (var i = 0 ; i < (otarfulist.value.length - 1)+1 ; i++ ){
      g_ssortstrings[__getArrayIndex(g_ssortstrings,1,i)] = ((((((((" "+txtFunction+" " + ocurrentoccfunc.value.ObjDef().Name(g_nloc)) + " ("+txtChapter+" ") + vbStr(ncurrentdepth.value)) + ".") + vbStr(nmodellevel.value)) + ".") + vbStr(g_nprozfunclevel)) + ")");
    }

    // Sorting of the functions in the list. Evaluation of the functions by the smallest X coordinate.
    // The others are written into the global list g_oepkfuncobjects.
    sortposition(onextobjocclist, nobjabspos);
    // Back into list oTarFuList.
    while (otarfulist.value.length > 0) {
      otarfulist.value = __delete(otarfulist, 0);
    }

    for (var i = 0 ; i < __getUBound(onextobjocclist, 1)+1 ; i++ ){
      otarfulist.value[__arrlength(otarfulist)] = onextobjocclist[__getArrayIndex(onextobjocclist,1,i)];
    }

    if (otarfulist.value.length > 1) {
      for (var i = 1 ; i < (otarfulist.value.length - 1)+1 ; i++ ){
        if (! (g_oepkfuncobjects.length == 0)) {
          g_ssourceruleoffunc = __redimArray(g_ssourceruleoffunc, new __ArrayInfo(0, (__getUBound(g_ssourceruleoffunc, 1) + 1)), "", true); 
        }

        g_oepkfuncobjects[__arrlength(g_oepkfuncobjects)] = otarfulist.value[i];
        g_ssourceruleoffunc[__getArrayIndex(g_ssourceruleoffunc,1,__getUBound(g_ssourceruleoffunc, 1))] = g_ssortstrings[__getArrayIndex(g_ssortstrings,1,i)];
      }
    }

    bdonefunc = __toBool(checkobj(new __holder(otarfulist.value[0]), sidofevaluatefun, ncurrentdepth, nmodellevel, new __holder(__toString((g_nprozfunclevel + 1))), false, true, g_odonemodfuncoccs, g_sdfuoccid, ""));
    if (bdonefunc) {
      g_bdonefunc = __toBool(true);
      g_sdonefuncid = __toString(sidofevaluatefun.value);
    }

    outfuncdata(ocurrentoccfunc, ndepth, ncurrentdepth, nmodellevel, new __holder(vbStr(g_nprozfunclevel)), oasprozmodellist, ssourcefuncproc, srulestring);
    g_bdonefunc = __toBool(false);
    if (! (bdonefunc)) {
      outofepkfunc(ocurrentmodel, new __holder(otarfulist.value[0]), ndepth, ncurrentdepth, nmodellevel, oasprozmodellist, ssourcefuncproc, ocurrentoccfunc.value.ObjDef().Name(g_nloc), false);
    }
  }

  if (otarfulist.value.length == 0 || bdonefunc) {
    outfuncdata(ocurrentoccfunc, ndepth, ncurrentdepth, nmodellevel, new __holder(vbStr(g_nprozfunclevel)), oasprozmodellist, ssourcefuncproc, srulestring);
    g_ooutfile.value.OutputLn(txtEndOfPath, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
    g_ooutfile.value.OutputLn("", getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
  }
}

// --------------------------------------------------------------
// Subprogram OutOfEPKFunc for the topological output of the functions in the model (eEPC).
// oCurrentModel = Current model.
// oCurrentOccFunc = current function (Occ).
// nDepth = Depth of the model level to be evaluated.
// nCurrentDepth = Current depth of the model levels.
// nModelLevel = Number or quantity of models that have already been output.
// oAsProzModelList = List of assigned models (of the Processes type).
// sSourceFuncProc =  List containing names of functions to which processes are assigned.
// sPredFuncOcc = 'Name of the predecessor function.
// bIsIndex = sPredFuncOcc is index (true) or name (false)

function outofepkfunc_std(ocurrentmodel, ocurrentoccfunc, ndepth, ncurrentdepth, nmodellevel, oasprozmodellist, ssourcefuncproc, spredfuncocc, bisindex)
{
  var ndummy = new __holder(0);
  var sdummy = new __holder("");

  var otargetfuncobjectlist = new __holder(null); 
  var srulestring = new __holder(""); 
  var srule = ""; 
  var objectxposition = __createArray(0); 
  var objectyposition = __createArray(0); 
  var bdonefunc = false; 
  var sidofevaluatefun = new __holder(""); 

  // Default settings
  var lnumberoftargetfuncobjects = 0;
  var onextobjocclist = __redimArray(onextobjocclist, new __ArrayInfo(0, 0), null, false); 
  otargetfuncobjectlist.value = new Array();

  switch(spredfuncocc) {
    case "":
      srulestring.value = __toString("");
    break;
    default:
      if (vbIsNumeric(spredfuncocc) && bisindex) {
        srulestring.value = __toString(g_ssourceruleoffunc[__getArrayIndex(g_ssourceruleoffunc,1,__toInteger(spredfuncocc))]);
      } else {
        srulestring.value = __toString(((((((((" "+txtFunction + spredfuncocc) + " ( " + txtChapter ) + vbStr(ncurrentdepth.value)) + ".") + vbStr(nmodellevel.value)) + ".") + vbStr(g_nprozfunclevel)) + ")"));
      }
  }

  g_nprozfunclevel = __toInteger((g_nprozfunclevel + 1));
  findnextfunc(ocurrentmodel, ocurrentoccfunc.value, otargetfuncobjectlist);
  bdonefunc = __toBool(checkobj(ocurrentoccfunc, sdummy, ndummy, ndummy, sdummy, false, false, g_odonemodfuncoccs, g_sdfuoccid, ocurrentoccfunc.value.ObjDef().Name(g_nloc)));
  if (otargetfuncobjectlist.value.length > 0) {
    objectxposition = __redimArray(objectxposition, new __ArrayInfo(0, (otargetfuncobjectlist.value.length - 1)), 0, false); 
    objectyposition = __redimArray(objectyposition, new __ArrayInfo(0, (otargetfuncobjectlist.value.length - 1)), 0, false); 
    onextobjocclist = __redimArray(onextobjocclist, new __ArrayInfo(0, (otargetfuncobjectlist.value.length - 1)), null, false); 
    for (var j = 0 ; j < (otargetfuncobjectlist.value.length - 1)+1 ; j++ ){
      onextobjocclist[__getArrayIndex(onextobjocclist,1,j)] = otargetfuncobjectlist.value[j];
      objectxposition[__getArrayIndex(objectxposition,1,j)] = otargetfuncobjectlist.value[j].X();
      objectyposition[__getArrayIndex(objectyposition,1,j)] = otargetfuncobjectlist.value[j].Y();
    }

    // The paths are inserted in the global list to make sure that they are taken into account when sorting.
    g_ssortstrings = __redimArray(g_ssortstrings, new __ArrayInfo(0, (otargetfuncobjectlist.value.length - 1)), "", true); 
    for (var i = 0 ; i < (otargetfuncobjectlist.value.length - 1)+1 ; i++ ){
      g_ssortstrings[__getArrayIndex(g_ssortstrings,1,i)] = ((((((((getString("TEXT54") + ocurrentoccfunc.value.ObjDef().Name(g_nloc)) + getString("TEXT51")) + vbStr(ncurrentdepth.value)) + ".") + vbStr(nmodellevel.value)) + ".") + vbStr(g_nprozfunclevel)) + ")");
    }

    // Sorting of the functions in the list. Evaluation of the functions by the smallest X coordinate.
    // The others are inserted into the global list g_oEpkFuncObjects.
    sortposition(onextobjocclist, objectxposition, objectyposition);
    g_bsamefunction = __toBool(false);
    // Back into list oTargetFuncObjectList.
    while (otargetfuncobjectlist.value.length > 0) {
      otargetfuncobjectlist.value = __delete(otargetfuncobjectlist, 0);
    }

    for (var i = 0 ; i < __getUBound(onextobjocclist, 1)+1 ; i++ ){
      otargetfuncobjectlist.value[__arrlength(otargetfuncobjectlist)] = onextobjocclist[__getArrayIndex(onextobjocclist,1,i)];
    }

    if (otargetfuncobjectlist.value.length > 1) {
      for (var i = 1 ; i < (otargetfuncobjectlist.value.length - 1)+1 ; i++ ){
        if (! (g_oepkfuncobjects.length == 0)) {
          g_ssourceruleoffunc = __redimArray(g_ssourceruleoffunc, new __ArrayInfo(0, (__getUBound(g_ssourceruleoffunc, 1) + 1)), "", true); 
        }

        g_oepkfuncobjects[__arrlength(g_oepkfuncobjects)] = otargetfuncobjectlist.value[i];
        g_ssourceruleoffunc[__getArrayIndex(g_ssourceruleoffunc,1,__getUBound(g_ssourceruleoffunc, 1))] = g_ssortstrings[__getArrayIndex(g_ssortstrings,1,i)];
      }
    }

    bdonefunc = __toBool(checkobj(new __holder(otargetfuncobjectlist.value[0]), sidofevaluatefun, ncurrentdepth, nmodellevel, new __holder(__toString((g_nprozfunclevel + 1))), false, true, g_odonemodfuncoccs, g_sdfuoccid, ""));
    if (bdonefunc == true) {
      g_bdonefunc = __toBool(true);
      g_sdonefuncid = __toString(sidofevaluatefun.value);
    }

    
    outfuncdata(ocurrentoccfunc, ndepth, ncurrentdepth, nmodellevel, new __holder(vbStr(g_nprozfunclevel)), oasprozmodellist, ssourcefuncproc, srulestring);
    g_bdonefunc = __toBool(false);
    if (bdonefunc == false) {
      outofepkfunc(ocurrentmodel, new __holder(otargetfuncobjectlist.value[0]), ndepth, ncurrentdepth, nmodellevel, oasprozmodellist, ssourcefuncproc, ocurrentoccfunc.value.ObjDef().Name(g_nloc), false);
    }
  }

  if (otargetfuncobjectlist.value.length == 0 || bdonefunc)
  {
    outfuncdata(ocurrentoccfunc, ndepth, ncurrentdepth, nmodellevel, new __holder(vbStr(g_nprozfunclevel)), oasprozmodellist, ssourcefuncproc, srulestring);
    //g_ooutfile.value.OutputLn(txtEndOfPath, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
    g_ooutfile.value.OutputLn("", getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
  }
}

function outofrelationships(ocxnoccs, ocurrentfuncocc)
{
  if(g_nOptOutput==0) {
    outofrelationships_std(ocxnoccs, ocurrentfuncocc);
  } else if(g_nOptOutput == 1) {
    outofrelationships_iso(ocxnoccs, ocurrentfuncocc);
  }
}


// --------------------------------------------------------------
// OutOfRelationships for outputting the allowed relationships.
// oCxnOccs = List of the maintained attributes.
// oCurrentFuncOcc = Current function occurrence.

function outofrelationships_std(ocxnoccs, ocurrentfuncocc)
{
  var bColored = true;   // variable to change background color of table rows          

  var osubmodfunc = null;   // Object definition of the current function definition in the assigned model.
  var soutstr = ""; 
  
  var scurrentfuname = __toString(ocurrentfuncocc.value.ObjDef().Name(g_nloc));
  g_ooutfile.value.TableRow();
  g_ooutfile.value.TableCell(txtRelations + ": ", 100, getString("TEXT48"), 12, Constants.C_BLACK, getTableCellColor_Bk(false), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
  //g_ooutfile.value.TableCell("", 0, getString("TEXT48"), 12, Constants.C_BLACK, getTableCellColor_Bk(false), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
  
  // Relationships within the model.
  for (var i = 0 ; i < (ocxnoccs.value.length - 1)+1 ; i++ ){
    var ocurrentcxnocc = ocxnoccs.value[i];
    g_ooutfile.value.TableRow();
    if (ocurrentfuncocc.value.IsEqual(ocurrentcxnocc.SourceObjOcc()) == true) {
      g_ooutfile.value.TableCell(ocurrentcxnocc.Cxn().ActiveType(), 50, getString("TEXT48"), 12, Constants.C_BLACK, getTableCellColor_Bk(bColored), 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
      soutstr = __toString(ocurrentcxnocc.TargetObjOcc().ObjDef().Name(g_nloc));
      if (soutstr == "") {
        soutstr = __toString("("+txtNotNamed+")");
      }
      g_ooutfile.value.TableCell(soutstr, 50, getString("TEXT48"), 12, Constants.C_BLACK, getTableCellColor_Bk(bColored), 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    }
    else {
      g_ooutfile.value.TableCell(ocurrentcxnocc.Cxn().PassiveType(), 50, getString("TEXT48"), 12, Constants.C_BLACK, getTableCellColor_Bk(bColored), 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
      soutstr = __toString(ocurrentcxnocc.SourceObjOcc().ObjDef().Name(g_nloc));
      if (soutstr == "") {
        soutstr = __toString("("+txtNotNamed+")");
      }
      g_ooutfile.value.TableCell(soutstr, 50, getString("TEXT48"), 12, Constants.C_BLACK, getTableCellColor_Bk(bColored), 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    }
    bColored = !bColored; // Change background color
  }

  // Relationships in the assigned function allocation diagram.
  var oassignedmodels = ocurrentfuncocc.value.ObjDef().AssignedModels();
  oassignedmodels = __sort(oassignedmodels, Constants.AT_NAME, Constants.SORT_NONE, Constants.SORT_NONE, g_nloc);

  if (oassignedmodels.length > 0) {
    for (var h = 0 ; h < (oassignedmodels.length - 1)+1 ; h++ ){
      switch(oassignedmodels[h].OrgModelTypeNum()) {        // TANR 216764
        case Constants.MT_FUNC_ALLOC_DGM:
          if (! (ocurrentfuncocc.value.Model().IsEqual(oassignedmodels[h]))) {
          // Assigned model is not current model
//            osubmodfunc = oassignedmodels[h].ObjOccListFilter(scurrentfuname, g_nloc, Constants.OT_FUNC);
            osubmodfunc = ocurrentfuncocc.value.ObjDef().OccListInModel(oassignedmodels[h]);    // Call-ID 106790

            ocxnoccs.value = oassignedmodels[h].CxnOccList();
            for ( i = 0 ; i < (ocxnoccs.value.length - 1)+1 ; i++ ){
              var ocurrentcxnocc = ocxnoccs.value[i];
              g_ooutfile.value.TableRow();

              if (osubmodfunc[0].IsEqual(ocurrentcxnocc.SourceObjOcc()) == true) {
                g_ooutfile.value.TableCell(ocurrentcxnocc.Cxn().ActiveType(), 50, getString("TEXT48"), 12, Constants.C_BLACK, getTableCellColor_Bk(bColored), 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
                g_ooutfile.value.TableCell(ocurrentcxnocc.TargetObjOcc().ObjDef().Name(g_nloc), 50, getString("TEXT48"), 12, Constants.C_BLACK, getTableCellColor_Bk(bColored), 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
              }
              else {
                g_ooutfile.value.TableCell(ocurrentcxnocc.Cxn().PassiveType(), 50, getString("TEXT48"), 12, Constants.C_BLACK, getTableCellColor_Bk(bColored), 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
                g_ooutfile.value.TableCell(ocurrentcxnocc.SourceObjOcc().ObjDef().Name(g_nloc), 50, getString("TEXT48"), 12, Constants.C_BLACK, getTableCellColor_Bk(bColored), 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
              }
            }
          }
        break;
      }
      bColored = !bColored; // Change background color
    }
  }
}


// --------------------------------------------------------------
// Subprogram OutOfRelationships for outputting the allowed relationships.

function outofrelationships_iso(ocxnoccs, ofuncocc)
{
  var osubcxnoccs = new __holder(null);   // List of sub connections.
  var osubmodfunc = null;   // Object definition of the current function definition in the assigned model.
  var scurrentfuname = "";   // Name of the current function.
  var sinputcxntype = __createArray("");   // List of connection types for input objects.
  var soutputcxntype = __createArray("");   // List of connection types for output objects.
  var sorgcxntype = __createArray("");   // List of connection types for organizational objects.
  var sapplcxntype = __createArray("");   // List of connection types for application system objects.

  sinputcxntype   = new Array();
  soutputcxntype  = new Array();
  sorgcxntype     = new Array();
  sapplcxntype    = new Array();

  var oinputobjoccs     = new Array();  // List of input objects of the model.
  var ooutputobjoccs    = new Array();  // List of output objects of the model.
  var oorgobjoccs       = new Array();  // List of organizational objects of the model.
  var oapplobjoccs      = new Array();  // List of application system objects of the model.

  scurrentfuname = __toString(ofuncocc.value.ObjDef().Name(g_nloc));
  // Relationships in the model are put in the corresponding list.
  getrelations(ocxnoccs, oinputobjoccs, ooutputobjoccs, oorgobjoccs, oapplobjoccs, sinputcxntype, soutputcxntype, sorgcxntype, sapplcxntype);
  // Relationships in the assigned function allocation diagram.
  var oassignedmodels = ofuncocc.value.ObjDef().AssignedModels();   // List of models which are assigned to the function.
  oassignedmodels = __sort(oassignedmodels, Constants.AT_NAME, Constants.SORT_NONE, Constants.SORT_NONE, g_nloc);

  if (oassignedmodels.length > 0) {
    for (var h = 0 ; h < (oassignedmodels.length - 1)+1 ; h++ ){
      switch(oassignedmodels[h].OrgModelTypeNum()) {        // TANR 216764
        case Constants.MT_FUNC_ALLOC_DGM:
          if (! (ofuncocc.value.Model().IsEqual(oassignedmodels[h]))) {
          // Assigned model is not current model
//            osubmodfunc = oassignedmodels[h].ObjOccListFilter_40(scurrentfuname, Constants.OT_FUNC, g_nloc);
            osubmodfunc = ofuncocc.value.ObjDef().OccListInModel(oassignedmodels[h]);    // Call-ID 106790
            
            osubcxnoccs.value = oassignedmodels[h].CxnOccList();
            getrelations(osubcxnoccs, oinputobjoccs, ooutputobjoccs, oorgobjoccs, oapplobjoccs, sinputcxntype, soutputcxntype, sorgcxntype, sapplcxntype);
          }
        break;
      }
    }
  }

  if (oinputobjoccs.length > 0) {
    g_ooutfile.value.TableRow();
    g_ooutfile.value.TableCell(txtInput+": ", 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.value.TableCell("", 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    relationout(oinputobjoccs, sinputcxntype);
  }

  if (ooutputobjoccs.length > 0) {
    g_ooutfile.value.TableRow();
    g_ooutfile.value.TableCell(txtOutput + ": ", 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.value.TableCell("", 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    relationout(ooutputobjoccs, soutputcxntype);
  }


  if (oorgobjoccs.length > 0) {
    g_ooutfile.value.TableRow();
    g_ooutfile.value.TableCell(txtOrgResp+": ", 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.value.TableCell("", 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    relationout(oorgobjoccs, sorgcxntype);
  }

  if (oapplobjoccs.length > 0) {
    g_ooutfile.value.TableRow();
    g_ooutfile.value.TableCell(txtApplSysSupp+": ", 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.value.TableCell("", 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    relationout(oapplobjoccs, sapplcxntype);
  }
}


// --------------------------------------------------------------
// Subprogram GetRelations for determining the relationships.
function getrelations(ocxnoccs, oinputobjoccs, ooutputobjoccs, oorgobjoccs, oapplobjoccs, sinputcxntype, soutputcxntype, sorgcxntype, sapplcxntype)
{
  var bcheck = false; 
  var sdummylist = __createArray(""); 
  var sdummy = new __holder(""); 
  var ndummy = new __holder(0); 

  for (var i = 0 ; i < (ocxnoccs.value.length - 1)+1 ; i++ ){
    switch(ocxnoccs.value[i].Cxn().TypeNum()) {
      case Constants.CT_PROV_INP_FOR:
      case Constants.CT_IS_INP_FOR:
        oinputobjoccs[__arrlength(oinputobjoccs)] = ocxnoccs.value[i].SourceObjOcc();
        bcheck = __toBool(checkobj(new __holder(ocxnoccs.value[i].SourceObjOcc().ObjDef()), sdummy, ndummy, ndummy, sdummy, false, false, new __holder(g_oinputobjdefs), sdummylist, ocxnoccs.value[i].SourceObjOcc().ObjDef().Name(g_nloc)));
        if (! (bcheck)) {
          g_oinputobjdefs[__arrlength(g_oinputobjdefs)] = ocxnoccs.value[i].SourceObjOcc().ObjDef();
        }

        sinputcxntype[sinputcxntype.length] = ocxnoccs.value[i].Cxn().PassiveType();
      break;

      case Constants.CT_HAS_OUT:
      case Constants.CT_CRT_OUT_TO:
        ooutputobjoccs[__arrlength(ooutputobjoccs)] = ocxnoccs.value[i].TargetObjOcc();
        bcheck = __toBool(checkobj(new __holder(ocxnoccs.value[i].TargetObjOcc().ObjDef()), sdummy, ndummy, ndummy, sdummy, false, false, new __holder(g_ooutputobjdefs), sdummylist, ocxnoccs.value[i].TargetObjOcc().ObjDef().Name(g_nloc)));
        if (! (bcheck)) {
          g_ooutputobjdefs[__arrlength(g_ooutputobjdefs)] = ocxnoccs.value[i].TargetObjOcc().ObjDef();
        }

        soutputcxntype[soutputcxntype.length] = ocxnoccs.value[i].Cxn().ActiveType();
      break;
    }

    switch(ocxnoccs.value[i].SourceObjOcc().ObjDef().TypeNum()) {
      case Constants.OT_SYS_ORG_UNIT:
      case Constants.OT_SYS_ORG_UNIT_TYPE:
      case Constants.OT_ORG_UNIT:
      case Constants.OT_ORG_UNIT_TYPE:
      case Constants.OT_PERS:
      case Constants.OT_PERS_TYPE:
      case Constants.OT_POS:
      case Constants.OT_GRP:
        oorgobjoccs[__arrlength(oorgobjoccs)] = ocxnoccs.value[i].SourceObjOcc();
        bcheck = __toBool(checkobj(new __holder(ocxnoccs.value[i].SourceObjOcc().ObjDef()), sdummy, ndummy, ndummy, sdummy, false, false, new __holder(g_oorgobjdefs), sdummylist, ocxnoccs.value[i].SourceObjOcc().ObjDef().Name(g_nloc)));
        if (! (bcheck)) {
          g_oorgobjdefs[__arrlength(g_oorgobjdefs)] = ocxnoccs.value[i].SourceObjOcc().ObjDef();
        }

        sorgcxntype[sorgcxntype.length] = ocxnoccs.value[i].Cxn().PassiveType();
      break;


      case Constants.OT_APPL_SYS:
      case Constants.OT_APPL_SYS_CLS:
      case Constants.OT_APPL_SYS_TYPE:
      case Constants.OT_MOD:
      case Constants.OT_MOD_CLS:
      case Constants.OT_MOD_TYPE:
      case Constants.OT_DP_FUNC:
      case Constants.OT_DP_FUNC_CLS:
      case Constants.OT_DP_FUNC_TYPE:
        oapplobjoccs[__arrlength(oapplobjoccs)] = ocxnoccs.value[i].SourceObjOcc();
        bcheck = __toBool(checkobj(new __holder(ocxnoccs.value[i].SourceObjOcc().ObjDef()), sdummy, ndummy, ndummy, sdummy, false, false, new __holder(g_oapplobjdefs), sdummylist, ocxnoccs.value[i].SourceObjOcc().ObjDef().Name(g_nloc)));
        if (! (bcheck)) {
          g_oapplobjdefs[__arrlength(g_oapplobjdefs)] = ocxnoccs.value[i].SourceObjOcc().ObjDef();
        }
        sapplcxntype[sapplcxntype.length] = ocxnoccs.value[i].Cxn().PassiveType();
      break;
    }

    if (ocxnoccs.value[i].TargetObjOcc().ObjDef().TypeNum() == Constants.OT_LOC) {
      oorgobjoccs[__arrlength(oorgobjoccs)] = ocxnoccs.value[i].TargetObjOcc();
      bcheck = __toBool(checkobj(new __holder(ocxnoccs.value[i].TargetObjOcc().ObjDef()), sdummy, ndummy, ndummy, sdummy, false, false, new __holder(g_oorgobjdefs), sdummylist, ocxnoccs.value[i].TargetObjOcc().ObjDef().Name(g_nloc)));
      if (! (bcheck)) {
        g_oorgobjdefs[__arrlength(g_oorgobjdefs)] = ocxnoccs.value[i].TargetObjOcc().ObjDef();
      }
      sorgcxntype[sorgcxntype.length] = ocxnoccs.value[i].Cxn().ActiveType();
    }
  }
}


// --------------------------------------------------------------
// Subprogram OutOfTheTreeFunc for recursive output of the functions in the model (function tree) topologically.
// oCurrentModel = Current model.
// oCurrentOccFunc = current function (Occ).
// nDepth = Depth of the model level to be evaluated.
// nCurrentDepth = Current depth of the model levels.
// nModelLevel = Number or quantity of models that have already been output.
// oAsProzModelList = List of assigned models (of the Processes type).
// sSourceFuncProc =  List containing names of functions to which processes are assigned.
// bFirst = Checking for first call of procedure.
// sBoughIn = Numbering of the branch when outputting function trees.
// bFirstSubFunc = Checking whether first call of SubFunktion.

var txtFuncTree  = getString("TEXT55");
var txtEndOfPath = getString("TEXT56");

function outofthetreefunc(ocurrentmodel, ocurrentoccfunc, ndepth, ncurrentdepth, nmodellevel, sfunclevel, oasprozmodellist, ssourcefuncproc, bfirst, sboughin, bfirstsubfunc)
{
  var oobjectlist = __createArray(null); 
  var objectxposition = __createArray(0); 
  var objectyposition = __createArray(0); 
  var bcheck = false;   // Indicator flag if FuncOcc has already been evaluated in the model = True.
  var sboughout = new __holder(""); 
  var sdummy = new __holder(""); 
  var sdummylist = __createArray(""); 

  var ilistsize = 0;
  if (bfirst.value == true) {
    g_sselecttion = __toString("");
  } else {
    g_sselecttion = __toString(txtFuncTree);
  }

  bfirst.value = __toBool(false);
  bcheck = __toBool(checkobj(ocurrentoccfunc, sdummy, ncurrentdepth, nmodellevel, sfunclevel, false, false, g_odonemodfuncoccs, sdummylist, ""));
  outfuncdata(ocurrentoccfunc, ndepth, ncurrentdepth, nmodellevel, sfunclevel, oasprozmodellist, ssourcefuncproc, sboughin);
  if (bcheck == false) {
    g_odonemodfuncoccs.value[__arrlength(g_odonemodfuncoccs.value)] = ocurrentoccfunc.value;
    var onextobjoccs = ocurrentmodel.value.GetSuccNodes(ocurrentoccfunc.value);
    if (onextobjoccs.length > 0) {
      for (var i = 0 ; i < (onextobjoccs.length - 1)+1 ; i++ ){
        var ocurrentnextobjocc = onextobjoccs[i];
        // The sub functions are written into a list for sorting.
        oobjectlist = __redimArray(oobjectlist, new __ArrayInfo(0, (onextobjoccs.length - 1)), null, true); 
        objectxposition = __redimArray(objectxposition, new __ArrayInfo(0, (onextobjoccs.length - 1)), 0, true); 
        objectyposition = __redimArray(objectyposition, new __ArrayInfo(0, (onextobjoccs.length - 1)), 0, true); 
        oobjectlist[__getArrayIndex(oobjectlist,1,i)] = ocurrentnextobjocc;
        objectxposition[__getArrayIndex(objectxposition,1,i)] = ocurrentnextobjocc.X();
        objectyposition[__getArrayIndex(objectyposition,1,i)] = ocurrentnextobjocc.Y();
      }

      sortposition(oobjectlist, objectxposition, objectyposition);
      for (var i = 0 ; i < (onextobjoccs.length - 1)+1 ; i++ ){
        if (bfirstsubfunc.value == true) {
          sboughout.value = __toString(vbStr((i + 1)));
          bfirstsubfunc.value = __toBool(false);
        } else {
          sboughout.value = __toString(((sboughin.value + ".") + vbStr((i + 1))));
        }
        outofthetreefunc(ocurrentmodel, new __holder(oobjectlist[__getArrayIndex(oobjectlist,1,i)]), ndepth, ncurrentdepth, nmodellevel, new __holder(((sfunclevel.value + ".") + vbStr((i + 1)))), oasprozmodellist, ssourcefuncproc, bfirst, sboughout, bfirstsubfunc);
      }
    }
  }
  else {
    g_bsamefunction = __toBool(false);
  }
}



// --------------------------------------------------------------
// OutputFuncOfModels Call of procedure for the selected sort criterion.
// oCurrentModel = Current model.
// nDepth = Depth of the model level to be evaluated.
// nCurrentDepth = Current depth of the model levels.
// nModelLevel =  Model number on the current level.
// oAsProzModelList = List of assigned models (of the Processes type).
// sSourceFuncProc =  List containing names of functions to which processes are assigned.

function outputfuncofmodels(ocurrentmodel, ndepth, ncurrentdepth, nmodellevel, oasprozmodellist, ssourcefuncproc)
{
    if(ocurrentmodel.value.TypeNum() == Constants.MT_VAL_ADD_CHN_DGM)
        outtopo(ocurrentmodel, ndepth, ncurrentdepth, nmodellevel, oasprozmodellist, ssourcefuncproc);
    else
        outnum(ocurrentmodel.value, ndepth, ncurrentdepth, nmodellevel, oasprozmodellist, ssourcefuncproc);
        
    
  //switch(g_nselectedoption) {
  //  case 0:
  //  case 1:
  //    outsymalpha(ocurrentmodel.value, ndepth, ncurrentdepth, nmodellevel, oasprozmodellist, ssourcefuncproc);
  //  break;

  //  case 2:
  //    outtopo(ocurrentmodel, ndepth, ncurrentdepth, nmodellevel, oasprozmodellist, ssourcefuncproc);
  //  break;

  //  case 3:
  //    outnum(ocurrentmodel.value, ndepth, ncurrentdepth, nmodellevel, oasprozmodellist, ssourcefuncproc);
  //  break;
  //}
}


function outputmodeldata(ocurrentmodel, ncurrentdepth, nmodellevel, ssourcefunc, bfirstmodel)
{
  if(g_nOptOutput==0) {
    outputmodeldata_std(ocurrentmodel, ncurrentdepth, nmodellevel, ssourcefunc, bfirstmodel);
  } else if(g_nOptOutput==1) {
    outputmodeldata_iso(ocurrentmodel, ncurrentdepth, nmodellevel, ssourcefunc, bfirstmodel);
  }
}

// --------------------------------------------------------------
// Subprogram OutputModelData for outputting the model data.
// Parameter
// oCurrentModel = Current model.
// nCurrentDepth = Current depth of the model levels.
// nModelLevel = Model number in the hierarchy.
// sSourceFunc = Function to which this model is assigned.
// bFirstModel = First model on this level (Yes = True / No = False).

function outputmodeldata_std(ocurrentmodel, ncurrentdepth, nmodellevel, ssourcefunc, bfirstmodel)
{
  var ocurrentsuperiorobject = new __holder(null); 
  var soutstring = new __holder("");   // Output string.
  var nmodelzoom = 0.0; 
  var bfirst = false; 
  var bdummy = false; 

  var osuperiorobjectlist = new Array();

  bfirst = __toBool(true);

  //if (bfirstmodel == true) {
  //  if (ncurrentdepth.value == 1) {
  //    g_ooutfile.value.OutputLnF((vbStr(ncurrentdepth.value) + "."+txtTopLevel), "REPORT1");
  //  } else {
  //    g_ooutfile.value.OutputLnF((vbStr(ncurrentdepth.value) + "."+txtLevel), "REPORT1");
  //  }
  //}
  
  g_ooutfile.value.OutputLn("", getString("TEXT48"), 8, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
  g_ooutfile.value.OutputLn(((((((vbStr(ncurrentdepth.value) + ".") + vbStr(nmodellevel.value)) + " ") + ocurrentmodel.value.Name(g_nloc) + " "))), getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0);
  //g_ooutfile.value.OutputLnF(((((((vbStr(ncurrentdepth.value) + ".") + vbStr(nmodellevel.value)) + " ") + ocurrentmodel.value.Name(g_nloc) + " "))), "REPORT2");
  g_ooutfile.value.OutputLn("", getString("TEXT48"), 8, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0);
  
  if(ocurrentmodel.value.TypeNum() == Constants.MT_VAL_ADD_CHN_DGM)
  {
  g_ooutfile.value.OutputLnF("", "REPORT1");
  g_ooutfile.value.BeginTable(100, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
    
  g_ooutfile.value.TableRow();
  g_ooutfile.value.TableCell("Model attributes:", 100, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
  //g_ooutfile.value.TableCell("", 0, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
  
  g_ooutfile.value.TableRow();
  g_ooutfile.value.TableCell("Type", 22, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
  g_ooutfile.value.TableCell(ocurrentmodel.value.Type(), 78, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
  
  g_ooutfile.value.TableRow();
  g_ooutfile.value.TableCell("Name", 22, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
  g_ooutfile.value.TableCell(ocurrentmodel.value.Attribute(Constants.AT_NAME, g_nloc).GetValue(true), 78, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
  
  g_ooutfile.value.TableRow();
  g_ooutfile.value.TableCell("Long Description", 22, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
  g_ooutfile.value.TableCell(ocurrentmodel.value.Attribute(Constants.AT_DESC, g_nloc).GetValue(false), 78, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    
  g_ooutfile.value.EndTable("", 100, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
  }
  
  if (g_bGraphic) {
    // Graphic of the model.
    graphicout(g_ooutfile, ocurrentmodel.value);
    g_ooutfile.value.OutputLn("", getString("TEXT48"), 8, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
    g_ooutfile.value.OutputLnF((((txtPicture + " " + vbStr(g_npicnum)) + "   ") + ocurrentmodel.value.Name(g_nloc)), "REPORT4");
    g_ooutfile.value.OutputLn("", getString("TEXT48"), 8, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
    g_npicnum++;
  }

}



// --------------------------------------------------------------
// Subprogram OutputModelData for outputting the model data.

function outputmodeldata_iso(ocurrentmodel, ncurrentdepth, nmodellevel, ssourcefunc, bfirstmodel)
{
  var ocurrentsuperiorobject = new __holder(null); 
  var ostrobjdef = null;   // List of structuring elements.
  var nmodelzoom = 0.0; 
  var soutstring = new __holder("");   // Output string.
  var bfirst = false; 
  var bdummy = false; 

  var osuperiorobjectlist = new Array();
  ostrobjdef = new Array();

  bfirst = __toBool(true);
  if (bfirstmodel) {
    if (ncurrentdepth.value == 1) {
      g_ooutfile.value.OutputLnF((vbStr(ncurrentdepth.value) + "."+txtTopLevel), "REPORT1");
    } else {
      g_ooutfile.value.OutputLnF((vbStr(ncurrentdepth.value) + "."+txtLevel), "REPORT1");
    }
  }

  osuperiorobjectlist = ocurrentmodel.value.SuperiorObjDefs();
  // Structuring elements are filtered.
  var i = 0;
  while (i <= (osuperiorobjectlist.length - 1)) {
    if (osuperiorobjectlist[i].TypeNum() == Constants.OT_STRCT_ELMT) {
      ostrobjdef[__arrlength(ostrobjdef)] = osuperiorobjectlist[i];
      osuperiorobjectlist = __delete(osuperiorobjectlist, i);
    }
    else {
      i = (i + 1);
    }
  }

  g_ooutfile.value.OutputLn("", getString("TEXT48"), 8, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
  g_ooutfile.value.OutputLnF(((((((vbStr(ncurrentdepth.value) + ".") + vbStr(nmodellevel.value)) + " ") + ocurrentmodel.value.Name(g_nloc)) + " ") + ocurrentmodel.value.Type()), "REPORT2");
  g_ooutfile.value.OutputLn("", getString("TEXT48"), 8, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
  g_ooutfile.value.BeginTable(100, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
  if (ssourcefunc.value == txtNull) {
    // No instruction.
  } else {
    g_ooutfile.value.TableRow();
    g_ooutfile.value.TableCell(txtAssignmentOf+": ", 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.value.TableCell(ssourcefunc.value, 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);

    if (osuperiorobjectlist.length > 0) {
      for ( i = 0 ; i < (osuperiorobjectlist.length - 1)+1 ; i++ ){
        ocurrentsuperiorobject.value = osuperiorobjectlist[i];
        bdummy = __toBool(checkobj(ocurrentsuperiorobject, soutstring, ncurrentdepth, nmodellevel, ssourcefunc, false, true, g_odonefuncdefs, g_sdfuoccid, ""));
        if (! (soutstring.value == "") && ! (soutstring.value == ssourcefunc.value)) {
          g_ooutfile.value.TableRow();
          if (bfirst) {
            g_ooutfile.value.TableCell(txtFurtherFunctionsWithSameAssignment, 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
            bfirst = __toBool(false);
          } else {
            g_ooutfile.value.TableCell("", 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
          }

          g_ooutfile.value.TableCell(soutstring.value, 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
          ocurrentsuperiorobject.value = null;
        }
      }
    }
  }

  g_ooutfile.value.TableRow();
  g_ooutfile.value.TableCell(txtGroup + ": ", 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
  g_ooutfile.value.TableCell(ocurrentmodel.value.Group().Name(g_nloc), 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
  var oattributes = ocurrentmodel.value.AttrList(g_nloc);
  if (oattributes.length > 0) {
    outofattributes(oattributes);
  }
  if (ostrobjdef.length > 0) {
    g_ooutfile.value.TableRow();
    g_ooutfile.value.TableCell(txtNormElements + ": ", 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.value.TableCell(txtNorm, 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    for ( i = 0 ; i < (ostrobjdef.length - 1)+1 ; i++ ){
      var ocurstrobjoccs = ostrobjdef[i].OccList();
      if (ocurstrobjoccs.length > 0) {
        soutstring.value = __toString(ocurstrobjoccs[0].Model().Name(g_nloc));
        for (var j = 1 ; j < (ocurstrobjoccs.length - 1)+1 ; j++ ){
          soutstring.value = __toString(((soutstring.value + ", ") + ocurstrobjoccs[j].Model().Name(g_nloc)));
        }
      }

      g_ooutfile.value.TableRow();
      g_ooutfile.value.TableCell(ostrobjdef[i].Name(g_nloc), 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
      g_ooutfile.value.TableCell(soutstring.value, 50, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
      soutstring.value = __toString("");
    }
  }

  g_ooutfile.value.EndTable("", 100, getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
  g_ooutfile.value.OutputLn("", getString("TEXT48"), 8, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0);
  if (g_bGraphic) {
    // Graphic of the model.
    graphicout(g_ooutfile, ocurrentmodel.value);

    g_ooutfile.value.OutputLn("", getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
    g_ooutfile.value.OutputLnF((((txtPicture + " " + vbStr(g_npicnum)) + vbChr(9)) + ocurrentmodel.value.Name(g_nloc)), "REPORT4");
    g_ooutfile.value.OutputLn("", getString("TEXT48"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
    g_npicnum++;
  }

  ocurrentsuperiorobject.value = null;
}


// --------------------------------------------------------------
// OutSymAlpha for outputting the functions in the model by symbol type or alphabetically.
// oCurrentModel = Current model.
// nDepth = Depth of the model level to be evaluated.
// nCurrentDepth = Current depth of the model levels.
// nModelLevel =  Model number on the current level.
// oAsProzModelList = List of assigned models (of the Processes type).
// sSourceFuncProc =  List containing names of functions to which processes are assigned.

function outsymalpha(ocurrentmodel, ndepth, ncurrentdepth, nmodellevel, oasprozmodellist, ssourcefuncproc)
{
  var ocurrentfuncocc = new __holder(null); 

  var ofuncoccs = ocurrentmodel.ObjOccListFilter(Constants.OT_FUNC);
  if (ofuncoccs.length > 0) {
    switch(g_nselectedoption) {
      case 0:
        ofuncoccs = __sort(ofuncoccs, Constants.SORT_SYMBOLNAME, Constants.AT_NAME, Constants.SORT_NONE, g_nloc);
      break;
      case 1:
        ofuncoccs = __sort(ofuncoccs, Constants.AT_NAME, Constants.SORT_NONE, Constants.SORT_NONE, g_nloc);
      break;
    }

    for (var i = 0 ; i < (ofuncoccs.length - 1)+1 ; i++ ){
      ocurrentfuncocc.value = ofuncoccs[i];
      outfuncdata(ocurrentfuncocc, ndepth, ncurrentdepth, nmodellevel, new __holder(vbStr((i + 1))), oasprozmodellist, ssourcefuncproc, new __holder(""));
      ocurrentfuncocc.value = null;
    }
  }
}


// --------------------------------------------------------------
// OutTopo for outputting the functions in the model after control flow.
// oCurrentModel = Current model.
// nDepth = Depth of the model level to be evaluated.
// nCurrentDepth = Current depth of the model levels.
// nModelLevel = Number or quantity of models that have already been output.
// oAsProzModelList = List of assigned models (of the Processes type).
// sSourceFuncProc =  List containing names of functions to which processes are assigned.

function outtopo(ocurrentmodel, ndepth, ncurrentdepth, nmodellevel, oasprozmodellist, ssourcefuncproc)
{
  var orootlist = new __holder(null);   // List of root objects.
  orootlist.value = new Array();

  findrootfunc(ocurrentmodel, orootlist);
  if (orootlist.value.length > 0) {
    switch(ocurrentmodel.value.OrgModelTypeNum()) {         // TANR 216764
      case Constants.MT_PRCS_SLCT_MTX:
        processselectionmatrixout(ocurrentmodel.value, ndepth, ncurrentdepth, nmodellevel, oasprozmodellist, ssourcefuncproc, orootlist.value);
      break;

      case Constants.MT_FUNC_TREE:
        functiontreeout(ocurrentmodel, ndepth, ncurrentdepth, nmodellevel, oasprozmodellist, ssourcefuncproc, orootlist.value);
      break;

      case Constants.MT_EEPC:
      case Constants.MT_EEPC_MAT:
      case Constants.MT_IND_PROC:
      case Constants.MT_OFFICE_PROC:
      case Constants.MT_PRCS_CHN_DGM:
      case Constants.MT_PCD_MAT:
      case Constants.MT_UML_ACTIVITY_DGM:
      case Constants.MT_EEPC_COLUMN:
      case Constants.MT_VAL_ADD_CHN_DGM:
      case Constants.MT_EEPC_ROW:
      case Constants.MT_EEPC_TAB:
      case Constants.MT_EEPC_TAB_HORIZONTAL:
        eepkout(ocurrentmodel, ndepth, ncurrentdepth, nmodellevel, oasprozmodellist, ssourcefuncproc, orootlist.value);
      break;
    }
  }
  orootlist.value = null;
}


// --------------------------------------------------------------
// ProcessSelectionmatrixOut for the topological output of the functions in the model (PAM) after control flow.
// oCurrentModel = Current model.
// nDepth = Depth of the model level to be evaluated.
// nCurrentDepth = Current depth of the model levels.
// nModelLevel = Number or quantity of models that have already been output.
// oAsProzModelList = List of assigned models (of the Processes type).
// sSourceFuncProc =  List containing names of functions to which processes are assigned.
// oRootList = List of root objects.

function processselectionmatrixout(ocurrentmodel, ndepth, ncurrentdepth, nmodellevel, oasprozmodellist, ssourcefuncproc, orootlist)
{
  var ocurrentrootobject = new __holder(null);

  var ocurrentnextobjocc = null;   // Current following element.
  var oobjectdummy = null; 

  var currentobjectname = ""; 
  var cxntargetobjectname = ""; 
  var cxnsourceobjectname = ""; 

  var nysize = __createArray(0); 
  var nydummy = 0; 
  var nsize = 0; 
  var bchange = false; 

  if (orootlist.length > 0) {
    for (var i = 0 ; i < (orootlist.length - 1)+1 ; i++ ){
      ocurrentrootobject.value = orootlist[i];
      outfuncdata(ocurrentrootobject, ndepth, ncurrentdepth, nmodellevel, new __holder(vbStr((i + 1))), oasprozmodellist, ssourcefuncproc, new __holder(""));
      var onextobjoccs = ocurrentmodel.GetSuccNodes(ocurrentrootobject.value);
      if (onextobjoccs.length > 0) {
        var oobjectocclist = __redimArray(oobjectocclist, new __ArrayInfo(0, (onextobjoccs.length - 1)), null, true); 
        nysize = __redimArray(nysize, new __ArrayInfo(0, (onextobjoccs.length - 1)), 0, true); 
        for (var j = 0 ; j < (onextobjoccs.length - 1)+1 ; j++ ){
          oobjectocclist[__getArrayIndex(oobjectocclist,1,j)] = onextobjoccs[j];
          nysize[__getArrayIndex(nysize,1,j)] = onextobjoccs[j].y();
        }

        // Sorting by Y-coordinate.
        bchange = __toBool(false);
        if (nsize > 0) {
          while (bchange == false) {
            bchange = __toBool(true);
            for (var j = 0 ; j < (onextobjoccs.length - 2)+1 ; j++ ){
              if (nysize[__getArrayIndex(nysize,1,j)] > nysize[__getArrayIndex(nysize,1,(j + 1))]) {
                nydummy = __toInteger(nysize[__getArrayIndex(nysize,1,j)]);
                nysize[__getArrayIndex(nysize,1,j)] = nysize[__getArrayIndex(nysize,1,(j + 1))];
                nysize[__getArrayIndex(nysize,1,(j + 1))] = nydummy;
                oobjectdummy = oobjectocclist[__getArrayIndex(oobjectocclist,1,j)];
                oobjectocclist[__getArrayIndex(oobjectocclist,1,j)] = oobjectocclist[__getArrayIndex(oobjectocclist,1,(j + 1))];
                oobjectdummy = oobjectocclist[__getArrayIndex(oobjectocclist,1,(j + 1))];
                bchange = __toBool(false);
              }
            }
          }

          // Output of the sorted functions.
          for (var j = 0 ; j < (onextobjoccs.length - 1)+1 ; j++ ){
            outfuncdata(new __holder(oobjectocclist[__getArrayIndex(oobjectocclist,1,j)]), ndepth, ncurrentdepth, nmodellevel, new __holder(((vbStr((i + 1)) + ".") + vbStr((j + 1)))), oasprozmodellist, ssourcefuncproc, new __holder(""));
          }
          oobjectocclist = __redimArray(oobjectocclist, new __ArrayInfo(0, 0), null, false); 
          nysize = __redimArray(nysize, new __ArrayInfo(0, 0), 0, false); 
          nsize = __toInteger(0);
        }
      }
      ocurrentrootobject.value = null;
    }
  }
}



// --------------------------------------------------------------
function RELATION(p_sCxnName, p_sObjName) {
    this.sCxnName = p_sCxnName;
    this.sObjName = p_sObjName;    
}

function sortRelation(a,b) {
    var tmp_lhs = new java.lang.String(a.sObjName);
    return tmp_lhs.compareTo(new java.lang.String(b.sObjName));
}

// Subprogram RelationOut for outputting the relationships.
function relationout(oobjoccs, scxntypename)
{ 
  // Anubis 274488    
  var aRelations = new Array();    
  for (var i = 0 ; i < oobjoccs.length; i++) {
      aRelations.push(new RELATION(scxntypename[__getArrayIndex(scxntypename,1,i)], oobjoccs[i].ObjDef().Name(g_nloc)));
  }
  aRelations.sort(sortRelation);
  
  var bColored = true;   // variable to change background color of table rows              
  for (var i = 0 ; i < aRelations.length; i++) {
    g_ooutfile.value.TableRow();
    g_ooutfile.value.TableCell(aRelations[i].sCxnName, 50, getString("TEXT48"), 12, Constants.C_BLACK, getTableCellColor_Bk(bColored), 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    g_ooutfile.value.TableCell(aRelations[i].sObjName, 50, getString("TEXT48"), 12, Constants.C_BLACK, getTableCellColor_Bk(bColored), 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);

    bColored = !bColored; // Change background color 
  }
}

// --------------------------------------------------------------
// SortPosition for sorting the objects, 1. by the x-Position 2. by the y- Position.
// oObjectOccList = List of functions which are sorted (Occ).
// ObjectXPosition = X-coordinates of the objects.
// ObjectYPosition = Y-coordinates of the objects.

function sortposition(oobjectocclist, objectxposition, objectyposition)
{
  var objectdummy = null; 
  var bchange = false; 
  var iydummy = 0; 
  var sdummy = ""; 

  // Sort by X-coordinate.
  bchange = __toBool(false);
  while (bchange == false) {
    bchange = __toBool(true);
    for (var j = 0 ; j < (__getUBound(oobjectocclist, 1) - 1)+1 ; j++ ){
      if (objectxposition[__getArrayIndex(objectxposition,1,j)] > objectxposition[__getArrayIndex(objectxposition,1,(j + 1))]) {
        iydummy = __toInteger(objectxposition[__getArrayIndex(objectxposition,1,j)]);
        objectxposition[__getArrayIndex(objectxposition,1,j)] = objectxposition[__getArrayIndex(objectxposition,1,(j + 1))];
        objectxposition[__getArrayIndex(objectxposition,1,(j + 1))] = iydummy;
        objectdummy = oobjectocclist[__getArrayIndex(oobjectocclist,1,j)];
        oobjectocclist[__getArrayIndex(oobjectocclist,1,j)] = oobjectocclist[__getArrayIndex(oobjectocclist,1,(j + 1))];
        oobjectocclist[__getArrayIndex(oobjectocclist,1,(j + 1))] = objectdummy;
        if (g_sselecttion == txtMTeEPC) {
          sdummy = __toString(g_ssortstrings[__getArrayIndex(g_ssortstrings,1,j)]);
          g_ssortstrings[__getArrayIndex(g_ssortstrings,1,j)] = g_ssortstrings[__getArrayIndex(g_ssortstrings,1,(j + 1))];
          g_ssortstrings[__getArrayIndex(g_ssortstrings,1,(j + 1))] = sdummy;
        }
        bchange = __toBool(false);
      }
    }
  }

  // Sort by y-coordinate.
  bchange = __toBool(false);
  while (bchange == false) {
    bchange = __toBool(true);
    for (var j = 0 ; j < (__getUBound(oobjectocclist, 1) - 1)+1 ; j++ ){
      if (objectxposition[__getArrayIndex(objectxposition,1,j)] == objectxposition[__getArrayIndex(objectxposition,1,(j + 1))]) {
        if (objectyposition[__getArrayIndex(objectyposition,1,j)] > objectyposition[__getArrayIndex(objectyposition,1,(j + 1))]) {
          iydummy = __toInteger(objectyposition[__getArrayIndex(objectyposition,1,j)]);
          objectyposition[__getArrayIndex(objectyposition,1,j)] = objectyposition[__getArrayIndex(objectyposition,1,(j + 1))];
          objectyposition[__getArrayIndex(objectyposition,1,(j + 1))] = iydummy;
          objectdummy = oobjectocclist[__getArrayIndex(oobjectocclist,1,j)];
          oobjectocclist[__getArrayIndex(oobjectocclist,1,j)] = oobjectocclist[__getArrayIndex(oobjectocclist,1,(j + 1))];
          oobjectocclist[__getArrayIndex(oobjectocclist,1,(j + 1))] = objectdummy;
          if (g_sselecttion == txtMTeEPC) {
            sdummy = __toString(g_ssortstrings[__getArrayIndex(g_ssortstrings,1,j)]);
            g_ssortstrings[__getArrayIndex(g_ssortstrings,1,j)] = g_ssortstrings[__getArrayIndex(g_ssortstrings,1,(j + 1))];
            g_ssortstrings[__getArrayIndex(g_ssortstrings,1,(j + 1))] = sdummy;
          }
          bchange = __toBool(false);
        }
      }
    }
  }
}


// --------------------------------------------------------------
// StringCut for cutting a string by a certain sign into a left and a right part.
// sDefaultString =  String that is cut by using a certain character.
// sSearchLeftString = left part of string.
// sSearchLeftString = right part of string.
// sSearchChar = Character by help o which the string is cut.

function stringcut(sdefaultstring, ssearchleftstring, ssearchrightstring, ssearchchar)
{
  var lpos = __toLong(vbInStr(sdefaultstring, ssearchchar));
  var lsize = __toLong(vbLen(sdefaultstring));
  if (! (lpos == 0)) {
    ssearchrightstring.value = __toString(vbRight(sdefaultstring, (lsize - lpos)));
    ssearchleftstring.value = __toString(vbLeft(sdefaultstring, (lpos - 1)));
  }
}




var bShowGraphicSettingsDialog;
var dlgFuncOutput;


/** 
  *  Dialog function for output options dialog
  * @param dlgItem   dialog item name
  * @param act action
  * @param suppVal support value
  * @return bool value, for act!=1, true means continue dialog and false ends dialog, vice versa for act=1 
  */
function dlgFuncOutputOptions(dlgItem, act, suppVal)
{
  var bResult = true;
  switch(act)
  {
    case 1:
      bResult = false;
      break;

    case 2:
      if(dlgItem==dicOutputFormat) {
        var bEnable = dlgFuncOutput.getDlgValue(dicOutputFormat)!==0;
        dlgFuncOutput.setDlgEnable(dicSupplement, bEnable);
      } else if(dlgItem==dicGraphic) {
        var bEnable = dlgFuncOutput.getDlgValue(dicGraphic)!=0;
        dlgFuncOutput.setDlgEnable(dicShowGraphicSettings, bEnable);
      } else if(dlgItem==dicShowGraphicSettings) {
        bShowGraphicSettingsDialog = true;
        bResult = false;
      }
      else if(dlgItem=="OK")
        bResult = false;
      else if(dlgItem=="Cancel")
        bResult = false;
      break;
  }

  return bResult;
}


// dialog item code constants
var dicSupplement       = "chkSuppl";
var dicOutputFormat     = "optOutput";
var dicSortOrder        = "optSortOrder";
var dicLinkLevels       = "txtLinkLevels";
var dicGraphic          = "chkGraphic";
var dicShowGraphicSettings  = "butShowGraphicSettings";

var txtModelTypes = new Array();
var dicModelTypes = new Array();

/**
 *  function showOutputOptionsDialog
 *  shows output options dialog with specified initial settings
 *  @param outfile output file
 *  @param holder_nOptOutputFormat receives output format setting
 *  @param holder_nOptSortOrder receives sort order settings
 *  @param holder_nLinkLevels receives link levels settings
 *  @param holder_aModelTypes receives model types settings
 *  @param holder_bGraphic receives graphic settings
 *  @return dialog return value
 */
function showOutputOptionsDialog(outfile, holder_nOptOutputFormat, holder_bSupplement, holder_nOptSortOrder, 
                              holder_nLinkLevels, holder_aModelTypes, holder_bGraphic)
{
  var ofilter = ArisData.getActiveDatabase().ActiveFilter();

  var userdialog = Dialogs.createNewDialogTemplate(0, 0, 610, 80 + (g_nmodeltypes.length*15) + 35, txtOutputOptionsDialogTitle, "dlgFuncOutputOptions");

  //userdialog.GroupBox(7, 0, 590, 65, txtOutputFormat);
  //userdialog.OptionGroup(dicOutputFormat);
  //userdialog.OptionButton(20, 15, 560, 15, txtOFModelHier);
  //userdialog.OptionButton(20, 30, 560, 15, txtOFModelHierISO);
  //userdialog.CheckBox(40, 45, 200, 15, txtSupplement, dicSupplement);

  //userdialog.GroupBox(7, 68, 590, 82, txtSortOrder);
  //userdialog.OptionGroup(dicSortOrder);
  //userdialog.OptionButton(20, 83, 570, 15, txtSortSymbolTypes);
  //userdialog.OptionButton(20, 98, 570, 15, txtSortAlpha);
  //userdialog.OptionButton(20, 113, 570, 15, txtSortTopologically);
  //userdialog.OptionButton(20, 128, 570, 15, txtSortNumerically);
  
  userdialog.GroupBox(7, 0, 590, 45, txtLinkLevels);
  userdialog.Text(20, 20, 140, 15, txtLinkLevels);
  userdialog.TextBox(185, 18, 60, 21, dicLinkLevels);

  userdialog.GroupBox(7, 50, 590, (g_nmodeltypes.length*15)+10, txtModelTypes);

  var y = 55; 
  for(var i =0; i<g_nmodeltypes.length;i++) {
    dicModelTypes[i] = "dicMT_" + g_nmodeltypes[i];
    txtModelTypes[i] = ofilter.ModelTypeName(g_nmodeltypes[i]);
    userdialog.CheckBox(20, y, 380, 15, txtModelTypes[i], dicModelTypes[i]);
    y += 15;
  }
  y += 15;

  userdialog.CheckBox(7, y, 180, 15, txtGraphic, dicGraphic);
  userdialog.PushButton(220, y, 150, 15, txtFormatGraphic, dicShowGraphicSettings);

  y+= 25;

  userdialog.OKButton();
  userdialog.CancelButton();
  userdialog.HelpButton("HID_a8f194f0_eae3_11d8_12e0_9d2843560f51_dlg_01.hlp");

  dlgFuncOutput = Dialogs.createUserDialog(userdialog); 
  
  // Read dialog settings from config  
  var sSection = "SCRIPT_28363b20-a2a1-11dd-7812-0017a4cd8aa8";
  ReadSettingsDlgValue(dlgFuncOutput, sSection, dicOutputFormat, holder_nOptOutputFormat.value);
  ReadSettingsDlgValue(dlgFuncOutput, sSection, dicSortOrder,    holder_nOptSortOrder.value);
  ReadSettingsDlgValue(dlgFuncOutput, sSection, dicSupplement,   holder_bSupplement.value?1:0);
  ReadSettingsDlgText(dlgFuncOutput, sSection, dicLinkLevels,   __toString(holder_nLinkLevels.value));

  for(var i=0;i<dicModelTypes.length;i++) {
      ReadSettingsDlgValue(dlgFuncOutput, sSection, dicModelTypes[i], 0);
  }
  ReadSettingsDlgValue(dlgFuncOutput, sSection, dicGraphic, holder_bGraphic.value?1:0);

  //dlgFuncOutput.setDlgEnable(dicSupplement, dlgFuncOutput.getDlgValue(dicOutputFormat) != 0);
  dlgFuncOutput.setDlgEnable(dicShowGraphicSettings, dlgFuncOutput.getDlgValue(dicGraphic) != 0);

  for(;;)
  {
    bShowGraphicSettingsDialog = false;
    nuserdialog = Dialogs.show( __currentDialog = dlgFuncOutput);
    // Displays dialog and waits for the confirmation with OK.
    if (nuserdialog == 0) {
      return nuserdialog;
    }
    if(bShowGraphicSettingsDialog) {
      showGraphicSettingsDialog(outfile);
      bShowGraphicSettingsDialog = false;
      continue;
    }
    else {
      if ((vbIsNumeric(dlgFuncOutput.getDlgText(dicLinkLevels)) == true)) {
        var depth = __toInteger(dlgFuncOutput.getDlgText(dicLinkLevels));
        if (depth > 0) {
          holder_nLinkLevels.value = depth;
          break; 
          
        } else { 
          Dialogs.MsgBox(txtNumberToSmall, vbOKOnly, "ARIS Report");
        }
      } else {
        Dialogs.MsgBox(txtPleaseNumber, vbOKOnly, "ARIS Report");
      }
    }
  }

  // Write dialog settings to config
  if (nuserdialog != 0) {  
      //WriteSettingsDlgValue(dlgFuncOutput, sSection, dicOutputFormat);
      //WriteSettingsDlgValue(dlgFuncOutput, sSection, dicSortOrder);
      //WriteSettingsDlgValue(dlgFuncOutput, sSection, dicSupplement);
      WriteSettingsDlgText(dlgFuncOutput, sSection, dicLinkLevels);    
      
      for(var i=0;i<dicModelTypes.length;i++) {
          WriteSettingsDlgValue(dlgFuncOutput, sSection, dicModelTypes[i]);
      }
      WriteSettingsDlgValue(dlgFuncOutput, sSection, dicGraphic);      
  }
  
  // set flag for output format
  holder_nOptOutputFormat.value = 0;
  holder_nOptSortOrder.value    = 0;
  holder_bSupplement.value      = false;
  
  holder_aModelTypes.value      = new Array();
  for(var i=0;i<dicModelTypes.length;i++) {
    if(dlgFuncOutput.getDlgValue(dicModelTypes[i])!=0) {
      holder_aModelTypes.value[holder_aModelTypes.value.length] = i;
    }
  }


  holder_bGraphic.value         = dlgFuncOutput.getDlgValue(dicGraphic)!=0;
    
  return nuserdialog;  
}


/**
 *  show graphic settings dialog, depending on bremote flag
 *
 */
function showGraphicSettingsDialog(outfile)
{
  var bcheckuserdialog = new __holder(true);
  if (bremote) { // WebDesigner
    graphicdialogs_default(outfile, bcheckuserdialog);
  }
  else { // WindowsClient
    graphicdialogs(outfile, bcheckuserdialog);
  }
}


main();











