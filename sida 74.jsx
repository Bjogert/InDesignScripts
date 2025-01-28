#target "InDesign"

/**
 * LeftAlignNumbersWithFrameWidth.jsx
 *
 * 1) Checks all text frames in the document.
 * 2) Skips frames that have letters or punctuation—only digits + whitespace allowed.
 * 3) Removes old tabs/spaces, splits the text into chunks, inserts single tabs.
 * 4) Decides how many chunks => picks tabPositions + frameWidth from settings.
 * 5) Applies left-aligned tab stops + sets paragraph justification to LEFT_ALIGN.
 * 6) (Optional) Resizes the text frame to frameWidth if defined.
 * 7) Single undo step. Progress bar included.
 */

app.doScript(
    function main() {
        // -----------------------------------------
        // 1) PROGRESS BAR / PALETTE
        // -----------------------------------------
        var w = new Window("palette", "Aligning Numeric Frames", undefined, { closeButton: false });
        w.orientation = "column";
        w.alignChildren = ["fill", "center"];

        var bar = w.add("progressbar", undefined, 0, 100);
        bar.preferredSize.width = 300;
        var statusText = w.add("statictext", undefined, "Starting...");
        statusText.alignment = ["fill", "center"];

        w.show();

        try {
            // -----------------------------------------
            // 2) DOCUMENT CHECK
            // -----------------------------------------
            if (!app.documents.length) {
                alert("No document open!");
                w.close();
                return;
            }
            var doc = app.activeDocument;

            // -----------------------------------------
            // 3) DEFINE PER-CHUNK SETTINGS
            // -----------------------------------------
            // For each chunk count, define:
            //   - tabPositions: array of left-aligned tab stop positions (in mm)
            //   - frameWidth: optional new width (in mm) for the text frame, or null to skip resizing
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
                // Extend if needed, e.g. 5, 6, ...
            };

            // -----------------------------------------
            // 4) SPEED SETTINGS
            // -----------------------------------------
            // Turn off preflight
            doc.preflightOptions.preflightOff = true;

            // Remember old measurement units
            var oldH = doc.viewPreferences.horizontalMeasurementUnits;
            var oldV = doc.viewPreferences.verticalMeasurementUnits;

            // Switch to mm so we can match the above positions easily
            doc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.MILLIMETERS;
            doc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.MILLIMETERS;

            // Unlock & show all layers
            for (var i = 0; i < doc.layers.length; i++) {
                doc.layers[i].locked = false;
                doc.layers[i].visible = true;
            }

            // Turn off redraw
            app.scriptPreferences.enableRedraw = false;

            // -----------------------------------------
            // 5) LOOP ALL TEXT FRAMES
            // -----------------------------------------
            var frames = doc.textFrames;
            var totalFrames = frames.length;
            bar.maxvalue = totalFrames; // if you get an error, use bar.maxValue

            var processed = 0;
            var changed = 0;

            for (var f = 0; f < totalFrames; f++) {
                var tf = frames[f];
                var didChange = processFrame(tf, settingsByChunkCount);
                if (didChange) changed++;
                processed++;

                // Update progress
                bar.value = f + 1;
                statusText.text = "Processing frame " + (f + 1) + " of " + totalFrames;
                w.update();
            }

            // -----------------------------------------
            // 6) RESTORE & REPORT
            // -----------------------------------------
            app.scriptPreferences.enableRedraw = true;
            doc.preflightOptions.preflightOff = false;
            doc.viewPreferences.horizontalMeasurementUnits = oldH;
            doc.viewPreferences.verticalMeasurementUnits = oldV;

            w.close();

            alert(
                "Checked " + processed + " text frames.\n" +
                changed + " numeric frames were reformatted.\n" +
                "All changes in one undo step!"
            );

        } catch (err) {
            alert("Error: " + err.message);
        } finally {
            if (w.visible) w.close();
            app.scriptPreferences.enableRedraw = true;
        }
    },
    ScriptLanguage.JAVASCRIPT,
    [],
    // Single undo step
    UndoModes.FAST_ENTIRE_SCRIPT,
    "Left Align Numbers with Optional Frame Width"
);


/**
 * processFrame(textFrame, settingsByChunkCount)
 * 1) Checks if text is purely numeric (digits + whitespace). If not, return false.
 * 2) Clean up old tabs/spaces, split, rejoin with \t
 * 3) Determine chunk count => get tabPositions & frameWidth
 * 4) Apply left-aligned tab stops & left justification
 * 5) (Optional) resize text frame
 * Returns true if content changed, false otherwise
 */
function processFrame(textFrame, settingsByChunkCount) {
    if (!textFrame.isValid || !textFrame.parentStory) return false;

    var oldContent = textFrame.contents;
    if (!oldContent) return false;

    // A) Remove tabs & extra spaces
    var newContent = oldContent.replace(/\t+/g, " ");   // remove tabs
    newContent = newContent.replace(/\s+/g, " ");       // multiple spaces => single
    newContent = newContent.replace(/^\s+|\s+$/g, "");  // trim

    // B) Check if purely numeric (digits + space)
    //    If there's any letter or punctuation, skip.
    var isNumericOnly = /^[0-9 ]*$/.test(newContent);
    if (!isNumericOnly) {
        return false; // skip non-numeric frames
    }
    if (!newContent) {
        return false; // nothing left
    }

    // C) Split => chunks
    var chunks = newContent.split(" "); // e.g. ["1234", "5678"] => 2 chunks

    // D) Rebuild with single tabs
    newContent = chunks.join("\t");

    var contentChanged = (newContent !== oldContent);
    if (contentChanged) {
        textFrame.contents = newContent;
    }

    // E) Figure out which config to use, based on chunk count
    var chunkCount = chunks.length;
    var config = settingsByChunkCount[chunkCount];

    // If no config, we can skip tab stops or do a fallback
    if (!config) {
        // e.g., do nothing extra
        return contentChanged;
    }

    // F) Apply tab stops (left-aligned) + left justification
    var paragraphs = textFrame.parentStory.paragraphs;
    for (var i = 0; i < paragraphs.length; i++) {
        var para = paragraphs[i];
        // Clear old stops
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
        // Force left align
        para.justification = Justification.LEFT_ALIGN;
    }

    // G) Resize frame if config.frameWidth != null
    if (config.frameWidth != null) {
        var gb = textFrame.geometricBounds; // [y1, x1, y2, x2]
        var top = gb[0];
        var left = gb[1];
        var bottom = gb[2];
        var right = left + config.frameWidth;
        textFrame.geometricBounds = [top, left, bottom, right];
    }

    return contentChanged;
}
