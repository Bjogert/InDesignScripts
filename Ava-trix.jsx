#target "InDesign"

/*
    HandleSpecialCase-Selection.jsx

    A standalone script that applies your "handleSpecialCase" logic
    to any selected text frames in InDesign.

    1) Checks for a document and selection.
    2) Loops through selected text frames.
    3) Applies handleSpecialCase to each:
       - Resizes the frame to config.frameWidth (extending right).
       - Prepends a tab, replaces spaces with tabs.
       - Left-aligns paragraphs and applies specified tab stops.
    4) Single undo step.

    You can modify "specialCaseConfig" for your desired tab positions
    and frame width (in millimeters).
*/

app.doScript(
    function main() {
        // -------------------------------------------------
        // 1) INITIAL CHECKS
        // -------------------------------------------------
        if (!app.documents.length) {
            alert("No document open!");
            return;
        }
        var doc = app.activeDocument;

        if (!app.selection || app.selection.length === 0) {
            alert("Nothing is selected. Please select one or more text frames.");
            return;
        }

        // Optional: Switch measurement units to mm if needed
        var oldH = doc.viewPreferences.horizontalMeasurementUnits;
        var oldV = doc.viewPreferences.verticalMeasurementUnits;
        doc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.MILLIMETERS;
        doc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.MILLIMETERS;

        // -------------------------------------------------
        // 2) CONFIG FOR SPECIAL CASE
        // -------------------------------------------------
        // Example values—adjust as you like:
        var specialCaseConfig = {
            frameWidth: 74,           // extends the right edge to left + 74 mm
            tabPositions: [15, 30, 45] // or whatever positions you need
        };

        // -------------------------------------------------
        // 3) FUNCTION: handleSpecialCase
        // -------------------------------------------------
        function handleSpecialCase(textFrame, config) {
            // 1) Resize the frame
            var gb = textFrame.geometricBounds; // [y1, x1, y2, x2]
            var top    = gb[0];
            var left   = gb[1];
            var bottom = gb[2];
            var right  = left + config.frameWidth;
            textFrame.geometricBounds = [top, left, bottom, right];

            // 2) Prepend a tab, replace spaces with tabs
            var content = textFrame.contents.replace(/^\s+|\s+$/g, "");
            content = "\t" + content.replace(/\s+/g, "\t");

            // Update text if changed
            if (textFrame.contents !== content) {
                textFrame.contents = content;
            }

            // 3) Left-align paragraphs, apply tab stops
            var paragraphs = textFrame.paragraphs;
            for (var j = 0; j < paragraphs.length; j++) {
                paragraphs[j].justification = Justification.LEFT_ALIGN;
                resetTabStops(paragraphs[j], config.tabPositions);
            }
        }

        // Helper to reset tab stops
        function resetTabStops(paragraph, positions) {
            // Remove old stops
            while (paragraph.tabStops.length > 0) {
                paragraph.tabStops[0].remove();
            }
            // Add new stops (left-aligned)
            for (var i = 0; i < positions.length; i++) {
                paragraph.tabStops.add({
                    alignment: TabStopAlignment.LEFT_ALIGN,
                    position: positions[i]
                });
            }
        }

        // -------------------------------------------------
        // 4) PROCESS SELECTED TEXT FRAMES
        // -------------------------------------------------
        app.scriptPreferences.enableRedraw = false; // speed up

        var framesProcessed = 0;

        for (var s = 0; s < app.selection.length; s++) {
            var obj = app.selection[s];

            // We only handle TextFrame objects
            if (obj.constructor.name === "TextFrame") {
                handleSpecialCase(obj, specialCaseConfig);
                framesProcessed++;
            }
        }

        // -------------------------------------------------
        // 5) RESTORE SETTINGS & REPORT
        // -------------------------------------------------
        doc.viewPreferences.horizontalMeasurementUnits = oldH;
        doc.viewPreferences.verticalMeasurementUnits = oldV;
        app.scriptPreferences.enableRedraw = true;

        if (framesProcessed > 0) {
            alert("Handled special case for " + framesProcessed + " text frames.");
        } else {
            alert("No text frames processed. Make sure you selected valid text frames.");
        }
    },
    ScriptLanguage.JAVASCRIPT,
    [],
    // Single undo step
    UndoModes.FAST_ENTIRE_SCRIPT,
    "Handle Special Case (Selection Only)"
);
