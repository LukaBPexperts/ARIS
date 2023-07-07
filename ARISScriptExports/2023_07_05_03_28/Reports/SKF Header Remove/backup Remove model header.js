var g_HeaderGroupings = new Array;
var g_HeaderGroupingMembers = new Array;

function main() {
    //Put all selected models in array o_SelectedModels
    var o_SelectedModels = ArisData.getSelectedModels();
    if (o_SelectedModels.length == 0) { //Folder selected instead of Models
        g_SelectedGroups = ArisData.getSelectedGroups();
        for (var i = 0; i < g_SelectedGroups.length; i++) {
            var o_TempModels = g_SelectedGroups[i].ModelList(true);
            o_SelectedModels.push.apply(o_SelectedModels, o_TempModels);
        }
        o_SelectedModels = ArisData.Unique(o_SelectedModels);
    }

    //Add header grouping in selected models to g_HeaderGroupings

    for (var i = 0; i < o_SelectedModels.length; i++) {
        var o_ModelGroupings = o_SelectedModels[i].getGroupings();
        for (var j = 0; j < o_ModelGroupings.length; j++) {
            var o_GroupingMembers = o_ModelGroupings[j].getMembers();
            for (var k = 0; k < o_ModelGroupings.length; k++) {
                if (o_GroupingMembers.indexOf(o_ModelGroupings[k]) < 0) {
                    var o_GroupingMembersMaxY = new Array;
                    for (var k = 0; k < o_GroupingMembers.length; k++) {
                        var o_GroupingMembersY = []
                        try{
                            if (o_GroupingMembers[k].KindNum() == "30004") {

                                o_GroupingMembersY.push(o_GroupingMembers[k].Y());
                                o_GroupingMembersMaxY = Math.max(o_GroupingMembersY);
                            }
                        }
                        catch(err){
                            o_GroupingMembersMaxY = 500;
                        }
                    }
                    //if (o_GroupingMembers.length == 12 && o_GroupingMembersMaxY < 275) {
                    if (o_GroupingMembers.length >1 && o_GroupingMembersMaxY < 450) {
                        for (var k = 0; k < o_GroupingMembers.length; k++) {
                            g_HeaderGroupings.push(o_ModelGroupings[j]);
                            g_HeaderGroupingMembers.push.apply(g_HeaderGroupingMembers, o_GroupingMembers);
                        }
                    }
                }
            }
        }
    }
    g_HeaderGroupings = ArisData.Unique(g_HeaderGroupings);
    g_HeaderGroupingMembers = ArisData.Unique(g_HeaderGroupingMembers);

    //Remove all Header grouping objects
    for (var i = 0; i < g_HeaderGroupingMembers.length; i++) {
        g_HeaderGroupingMembers[i].Remove();
    }
    
    
}
main()

function getObjectOccSource(cxnOcc, newModel) {
    var sourceObjOcc = cxnOcc.getSource()
    var sourceObjOccX = sourceObjOcc.X()
    var sourceObjOccY = sourceObjOcc.Y()
    var sourceObjDef = sourceObjOcc.ObjDef()

    var occListModel = sourceObjDef.OccListInModel(newModel)

    for (var i = 0; i < occListModel.length; i++) {
        var tempX = occListModel[i].X()
        var tempY = occListModel[i].Y()

        if ((tempX == sourceObjOccX) && (tempY == sourceObjOccY)) {

            return occListModel[i]
        }
    }
}

function getObjectOccTarget(cxnOcc, newModel) {
    var targetObjOcc = cxnOcc.getTarget()
    var targetObjOccX = targetObjOcc.X()
    var targetObjOccY = targetObjOcc.Y()
    var targetObjDef = targetObjOcc.ObjDef()

    var occListModel = targetObjDef.OccListInModel(newModel)

    for (var i = 0; i < occListModel.length; i++) {
        var tempX = occListModel[i].X()
        var tempY = occListModel[i].Y()

        if ((tempX == targetObjOccX) && (tempY == targetObjOccY)) {

            return occListModel[i]
        }
    }
}