#target "InDesign"

(function() {
    app.doScript(main, ScriptLanguage.JAVASCRIPT, [], UndoModes.ENTIRE_SCRIPT, "Replace Script Labels");

    function main() {
        if (!app.documents.length) {
            alert("No document open.");
            return;
        }

        var doc = app.activeDocument;
        var initialLabel = "";

        var priceGroups = [
            "SEK_SE", "SEK_NO", "DKK_DK",
            "EUR_EU", "GBP_UK", "EUR_BEIR",
            "EUR_EURAEA", "USD_US", "EUR_DE"
        ];

        // Corrected suffix trimming for ExtendScript compatibility
        if (app.selection.length === 1 && app.selection[0] instanceof TextFrame) {
            initialLabel = app.selection[0].label || "";
            for (var p = 0; p < priceGroups.length; p++) {
                var suffix = "_" + priceGroups[p];
                if (initialLabel.length > suffix.length && initialLabel.substr(-suffix.length) === suffix) {
                    initialLabel = initialLabel.substr(0, initialLabel.length - suffix.length);
                    break;
                }
            }
        }

        var dlg = app.dialogs.add({name: "Replace Script Labels", canCancel: true});
        var col1 = dlg.dialogColumns.add();
        var col2 = dlg.dialogColumns.add();

        var rowOld = col1.dialogRows.add();
        rowOld.staticTexts.add({staticLabel: "Old base label:"});
        var oldLabelField = rowOld.textEditboxes.add({editContents: initialLabel, minWidth: 160});

        var rowNew = col1.dialogRows.add();
        rowNew.staticTexts.add({staticLabel: "New base label:"});
        var newLabelField = rowNew.textEditboxes.add({editContents: "", minWidth: 160});

        var rowHeading = col2.dialogRows.add();
        rowHeading.staticTexts.add({staticLabel: "Select Price Groups:"});

        var groupsPanel = col2.borderPanels.add();
        var gpCol = groupsPanel.dialogColumns.add();

        var cbxGroups = [];
        for (var p = 0; p < priceGroups.length; p++) {
            var r = gpCol.dialogRows.add();
            var cb = r.checkboxControls.add({staticLabel: priceGroups[p], checkedState: false});
            cbxGroups.push({name: priceGroups[p], checkbox: cb});
        }

        var rowAll = gpCol.dialogRows.add();
        var cbAll = rowAll.checkboxControls.add({staticLabel: "ALL price groups", checkedState: false});

        if (!dlg.show()) {
            dlg.destroy();
            return;
        }

        var oldBase = oldLabelField.editContents;
        var newBase = newLabelField.editContents;
        var allChecked = cbAll.checkedState;

        var selectedGroups = [];
        if (!allChecked) {
            for (var sg = 0; sg < cbxGroups.length; sg++) {
                if (cbxGroups[sg].checkbox.checkedState) {
                    selectedGroups.push(cbxGroups[sg].name);
                }
            }
        }
        dlg.destroy();

        if (!oldBase) {
            alert("No old base label specified.");
            return;
        }
        if (!newBase) {
            alert("No new base label specified.");
            return;
        }
        if (!allChecked && selectedGroups.length === 0) {
            alert("No price groups selected.");
            return;
        }

        var layers = doc.layers, originalStates = [];
        for (var i = 0; i < layers.length; i++) {
            originalStates[i] = {locked: layers[i].locked, visible: layers[i].visible};
            layers[i].locked = false;
            layers[i].visible = true;
        }

        var replacedCount = 0;
        var pagesAffected = {};
        var pages = doc.pages;
        
        if (allChecked) {
            var prefixToFind = oldBase + "_";
            for (var a = 0; a < pages.length; a++) {
                var framesA = pages[a].textFrames;
                for (var b = 0; b < framesA.length; b++) {
                    var currentLabel = framesA[b].label;
                    if (currentLabel.indexOf(prefixToFind) === 0) {
                        var suffix = currentLabel.substring(oldBase.length);
                        var newLabel = newBase + suffix;
                        if (framesA[b].label !== newLabel) {
                            framesA[b].label = newLabel;
                            replacedCount++;
                            pagesAffected[pages[a].name] = true;
                        }
                    }
                }
            }
        } else {
            for (var g = 0; g < selectedGroups.length; g++) {
                var oldFullLabel = oldBase + "_" + selectedGroups[g];
                var newFullLabel = newBase + "_" + selectedGroups[g];
        
                for (var p2 = 0; p2 < pages.length; p2++) {
                    var frames2 = pages[p2].textFrames;
                    for (var f2 = 0; f2 < frames2.length; f2++) {
                        if (frames2[f2].label === oldFullLabel) {
                            frames2[f2].label = newFullLabel;
                            replacedCount++;
                            pagesAffected[pages[p2].name] = true;
                        }
                    }
                }
            }
        }
        
        

        restoreLayerStates(layers, originalStates);

        var affectedPagesList = [];
        for (var pageName in pagesAffected) {
            if (pagesAffected.hasOwnProperty(pageName)) {
                affectedPagesList.push(pageName);
            }
        }
        affectedPagesList.sort(function(a,b){ return Number(a) - Number(b); });
        affectedPagesList = affectedPagesList.join(", ");
        
        if (replacedCount > 0) {
            alert("Replaced " + replacedCount + " label(s).\nPages affected: " + affectedPagesList);
        } else {
            alert("No matching labels were found.");
        }
        
        function restoreLayerStates(layers, states) {
            for (var i = 0; i < layers.length; i++) {
                layers[i].locked = states[i].locked;
                layers[i].visible = states[i].visible;
            }
    }
}
})();