#target "InDesign"

/**
 * LeftAlignNumbersWithFrameWidth-SelectionOnly.jsx
 *
 * 1) Processes only SELECTED text frames in the active InDesign document.
 * 2) Skips frames containing any letters/punctuation (only digits + whitespace allowed).
 * 3) Removes old tabs/spaces, splits the text into chunks, re-inserts single tabs.
 * 4) Chooses tab stops + optional frame width (extending to the right) based on chunk count.
 * 5) Applies left-aligned tab stops and paragraph justification (LEFT_ALIGN).
 * 6) Maintains a progress bar for the selected text frames.
 * 7) All changes happen in a single undo step (UNDO).
 *
 * Note: We have removed:
 *   - The code that unlocked and made all layers visible.
 *   - The code that turned off preflight.
 *
 * Usage:
 *   - Select one or more text frames on an unlocked, visible layer.
 *   - Run this script. Only numeric frames in the selection are processed.
 */

app.doScript(
    function main() {
        // -----------------------------------------
        // 1) BUILD PROGRESS PALETTE
        // -----------------------------------------
        var w = new Window("palette", "Aligning Numeric Frames (Selection)", undefined, { closeButton: false });
        w.orientation = "column";
        w.alignChildren = ["fill", "center"];

        var bar = w.add("progressbar", undefined, 0, 100);
        bar.preferredSize.width = 300;
        var statusText = w.add("statictext", undefined, "Starting...");
        statusText.alignment = ["fill", "center"];

        w.show();

        try {
            // -----------------------------------------
            // 2) CHECK DOCUMENT & SELECTION
            // -----------------------------------------
            if (!app.documents.length) {
                alert("No document open!");
                w.close();
                return;
            }
            var doc = app.activeDocument;

            // If no objects are selected, nothing to do
            if (!app.selection || app.selection.length === 0) {
                alert("No objects selected. Please select one or more text frames.");
                w.close();
                return;
            }

            // -----------------------------------------
            // 3) CHUNK-BASED SETTINGS
            // -----------------------------------------
            // For each chunk count, define:
            //   tabPositions: left-aligned tab stops (in mm)
            //   frameWidth: optional new width (in mm). If null, no resizing.
            var settingsByChunkCount = {
                1: {
                    tabPositions: [],
                    frameWidth: 10
                },
                2: {
                    tabPositions: [73],
                    frameWidth: 85
                },
                3: {
                    tabPositions: [25.7, 51],
                    frameWidth: 75
                },
                4: {
                    tabPositions: [21.6, 47.5, 72.7],
                    frameWidth: 100
                }
                // Extend if needed (e.g. 5, 6, etc.)
            };

            // -----------------------------------------
            // 4) OPTIONAL: CHANGE UNITS
            // -----------------------------------------
            // If you want to ensure tabPositions match mm, set measurement units:
            var oldH = doc.viewPreferences.horizontalMeasurementUnits;
            var oldV = doc.viewPreferences.verticalMeasurementUnits;
            doc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.MILLIMETERS;
            doc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.MILLIMETERS;

            // (We are NOT turning off preflight or unlocking layers now.)

            // Turn off redraw for speed
            app.scriptPreferences.enableRedraw = false;

            // -----------------------------------------
            // 5) FILTER SELECTED TEXT FRAMES
            // -----------------------------------------
            var selectedFrames = [];
            for (var s = 0; s < app.selection.length; s++) {
                var selObj = app.selection[s];
                // We only want text frames
                if (selObj.constructor.name === "TextFrame") {
                    selectedFrames.push(selObj);
                }
            }

            if (selectedFrames.length === 0) {
                alert("No text frames found in the selection.");
                w.close();
                return;
            }

            var totalFrames = selectedFrames.length;
            bar.maxvalue = totalFrames; // If error, try bar.maxValue

            var processed = 0;
            var changed = 0;

            // -----------------------------------------
            // 6) PROCESS SELECTED FRAMES
            // -----------------------------------------
            for (var i = 0; i < selectedFrames.length; i++) {
                var tf = selectedFrames[i];
                var didChange = processFrame(tf, settingsByChunkCount);
                if (didChange) changed++;
                processed++;

                // Update the progress bar & text
                bar.value = processed;
                statusText.text = "Processing " + processed + " / " + totalFrames + " text frames...";
                w.update();
            }

            // -----------------------------------------
            // 7) RESTORE & REPORT
            // -----------------------------------------
            // Re-enable redraw
            app.scriptPreferences.enableRedraw = true;

            // Restore measurement units
            doc.viewPreferences.horizontalMeasurementUnits = oldH;
            doc.viewPreferences.verticalMeasurementUnits = oldV;

            // Close progress window
            w.close();

            // Final message
            alert(
                "Scanned " + processed + " selected text frames.\n" +
                changed + " numeric frames were reformatted.\n" +
                "All changes in one undo step!"
            );

        } catch (err) {
            alert("Error: " + err.message);
        } finally {
            // Ensure window closes even on error
            if (w.visible) w.close();
            app.scriptPreferences.enableRedraw = true;
        }
    },
    ScriptLanguage.JAVASCRIPT,
    [],
    // Single undo step
    UndoModes.FAST_ENTIRE_SCRIPT,
    "Left Align Numbers (Selection Only)"
);


/**
 * processFrame(textFrame, settingsByChunkCount)
 * 1) Checks if content is purely numeric (digits + whitespace).
 * 2) Cleans up tabs/spaces, splits into chunks => rejoin with \t.
 * 3) Determines chunk count => picks tabPositions & frameWidth from config.
 * 4) Applies left-aligned tab stops & left justification to each paragraph.
 * 5) Optionally resizes the text frame (extend right edge).
 * Returns true if content changed, false otherwise.
 */
function processFrame(textFrame, settingsByChunkCount) {
    if (!textFrame.isValid || !textFrame.parentStory) {
        return false;
    }

    var oldContent = textFrame.contents;
    if (!oldContent) {
        return false;
    }

    // 1) Remove old tabs + multiple spaces, trim
    var newContent = oldContent.replace(/\t+/g, " ");
    newContent = newContent.replace(/\s+/g, " ");
    newContent = newContent.replace(/^\s+|\s+$/g, "");

    // 2) Must be digits + spaces only
    var isNumericOnly = /^[0-9 ]*$/.test(newContent);
    if (!isNumericOnly) {
        return false; // skip non-numeric
    }
    if (!newContent) {
        return false; // nothing left
    }

    // 3) Split => chunks, rejoin with tabs
    var chunks = newContent.split(" ");
    newContent = chunks.join("\t");

    var contentChanged = (newContent !== oldContent);
    if (contentChanged) {
        textFrame.contents = newContent;
    }

    // 4) Determine chunkCount => config
    var chunkCount = chunks.length;
    var config = settingsByChunkCount[chunkCount];
    if (!config) {
        // No config for this chunk count => do nothing else
        return contentChanged;
    }

    // Apply left-aligned tab stops & justification
    var paragraphs = textFrame.parentStory.paragraphs;
    for (var p = 0; p < paragraphs.length; p++) {
        var para = paragraphs[p];
        // Remove existing tab stops
        while (para.tabStops.length > 0) {
            para.tabStops[0].remove();
        }
        // Add new stops
        var positions = config.tabPositions || [];
        for (var t = 0; t < positions.length; t++) {
            var ts = para.tabStops.add();
            ts.alignment = TabStopAlignment.LEFT_ALIGN;
            ts.position = positions[t];
        }
        para.justification = Justification.LEFT_ALIGN;
    }

    // 5) Resize frame if config.frameWidth is set
    if (config.frameWidth != null) {
        var gb = textFrame.geometricBounds; // [y1, x1, y2, x2]
        var top  = gb[0];
        var left = gb[1];
        var bottom = gb[2];
        var right = left + config.frameWidth; // keep left edge, shift right edge
        textFrame.geometricBounds = [top, left, bottom, right];
    }

    return contentChanged;
}
