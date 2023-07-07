var selectedModels = ArisData.getSelectedModels()

//var objOccs = selectedModels[0].ObjOccListFilter(Constants.OT_BPMN_POOL)
var objOccs = selectedModels[0].ObjOccListFilter()

var designerComp = Context.getComponent("Designer")

var horizontalFčag = designerComp.getBPMNSupport(selectedModels[0]).isHorizontal (objOccs[0] )

var a=2