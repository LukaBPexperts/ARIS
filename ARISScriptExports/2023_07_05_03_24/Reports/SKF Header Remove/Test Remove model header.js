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

    for (var i = 0; i < o_SelectedModels.length; i++) {
        Context.writeStatus("Model: " + (i + 1) + '/' + (o_SelectedModels.length + 1) + "  " + o_SelectedModels[i].Name(Context.getSelectedLanguage()))

        var o_ModelGroupings = o_SelectedModels[i].getGroupings();
        for (var j = 0; j < o_ModelGroupings.length; j++) {
            var o_GroupingMembers = o_ModelGroupings[j].getMembers();
            if (o_GroupingMembers.indexOf(o_ModelGroupings[j]) < 0) {
                var o_GroupingMembersMaxY = new Array;
                for (var l = 0; l < o_GroupingMembers.length; l++) {
                    var o_GroupingMembersY = []
                    try {
                        if (o_GroupingMembers[l].KindNum() == "30004" || o_GroupingMembers[l].KindNum() == 30003) {
                            o_GroupingMembersY.push(o_GroupingMembers[l].Y());
                            o_GroupingMembersMaxY = Math.max(o_GroupingMembersY);
                        } else if (o_GroupingMembers[l].KindNum() == Constants.CID_GFXOBJ) {
                            o_GroupingMembersY.push(o_GroupingMembers[l].getY());
                            o_GroupingMembersMaxY = Math.max(o_GroupingMembersY);
                        } else {
                            o_GroupingMembersY.push(500);
                            o_GroupingMembersMaxY = Math.max(o_GroupingMembersY);
                        }
                    } catch (err) {
                        o_GroupingMembersMaxY = 500;
                    }
                }
                if (o_GroupingMembers.length > 2 && o_GroupingMembersMaxY < 450) {
                    Context.setProperty("maxY", o_GroupingMembersMaxY)
                    g_HeaderGroupingMembers.push.apply(g_HeaderGroupingMembers, o_GroupingMembers);
                }
            }
        }
    }
    g_HeaderGroupingMembers = ArisData.Unique(g_HeaderGroupingMembers);

    //Remove all Header grouping objects
    Context.setProperty("flagRemoved", "false")
    for (var i = 0; i < g_HeaderGroupingMembers.length; i++) {
        var flagRemoved = g_HeaderGroupingMembers[i].Remove();
        if (flagRemoved == true) {
            Context.setProperty("flagRemoved", "true")
        }
    }
}
main()