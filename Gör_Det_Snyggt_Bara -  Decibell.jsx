#target "InDesign"

/**
 * HandleNumbers-WithRangesAndWidths.jsx
 *
 * Demonstrates how to:
 *  - Identify frames by numeric content OR by frame width.
 *  - Apply different tab stops/widths.
 *  - Skip certain pages if desired.
 *  - Use a progress bar.
 *  - Wrap in one undo step for performance/single history entry.
 */

app.doScript(
    function main() {
        // -------------------------------------------------
        // PROGRESS WINDOW SETUP
        // -------------------------------------------------
        var progressWindow = new Window("palette", "Sprinkling Fairy Dust", undefined, { closeButton: false });
        progressWindow.orientation = "column";
        progressWindow.alignChildren = ["fill", "center"];

        var progressBar = progressWindow.add("progressbar", undefined, 0, 100);
        progressBar.preferredSize.width = 300;
        var progressText = progressWindow.add("statictext", undefined, "Gathering magical energy...");
        progressText.alignment = ["fill", "center"];

        progressWindow.show();

        try {
            // -------------------------------------------------
            // GET CURRENT DOCUMENT
            // -------------------------------------------------
            var doc = app.activeDocument;

            // -------------------------------------------------
            // SETTINGS (TAB POSITIONS & FRAME WIDTH)
            // -------------------------------------------------
            // 1) By numeric pattern approach (like your original script)
            // 2) By known frame width approach (for "identifying" certain boxes)
            // Adjust these as needed:
            var settings = {
                sixNumbers: {
                    tabPositions: [13.05, 26.1, 39.1, 51.86, 64.68],
                    frameWidth: 74
                },
                fourNumbers: {
                    tabPositions: [15, 30, 45],
                    frameWidth: 57
                },
                threeNumbers: {
                    tabPositions: [15, 30, 45],
                    frameWidth: 42
                },
                twoNumbers: {
                    tabPositions: [26],
                    frameWidth: 35
                },
                singleNumber: {
                    tabPositions: [0],
                    frameWidth: 10
                }
            };

            // -------------------------------------------------
            // OPTIONAL: SKIP SPECIFIC PAGES
            // -------------------------------------------------
            // Example: skip pages named "2" and "5"
            // (In InDesign, page.name is typically "1", "2", etc.
            //  or it might be alternate like "v", "xi", etc.)
            var skipThesePages = ["2", "5"];  // Example

            // -------------------------------------------------
            // UNLOCK & SHOW ALL LAYERS
            // -------------------------------------------------
            var layers = doc.layers;
            for (var i = 0; i < layers.length; i++) {
                layers[i].locked = false;
                layers[i].visible = true;
            }

            // -------------------------------------------------
            // SPEED SETTINGS (MEASUREMENT + PREFLIGHT)
            // -------------------------------------------------
            doc.preflightOptions.preflightOff = true;
            var originalHUnits = doc.viewPreferences.horizontalMeasurementUnits;
            var originalVUnits = doc.viewPreferences.verticalMeasurementUnits;
            doc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.MILLIMETERS;
            doc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.MILLIMETERS;
            app.scriptPreferences.enableRedraw = false;

            // -------------------------------------------------
            // PAGE / FRAME LOOP
            // -------------------------------------------------
            var pages = doc.pages;
            var totalPages = pages.length;
            var updatedCount = 0;  // # of frames that matched pattern or width
            var modifiedCount = 0; // # of frames whose content was actually changed

            // If your InDesign version requires .maxValue (capital V), change below:
            progressBar.maxvalue = totalPages;

            // Regex patterns for numeric matching:
            var matchSix = /^\d+\s+\d+\s+\d+\s+\d+\s+\d+\s+\d+$/;
            var matchFour = /^\d+\s+\d+\s+\d+\s+\d+$/;
            var matchThree = /^\d+\s+\d+\s+\d+$/;
            var matchTwo = /^\d+\s+\d+$/;
            var matchSingle = /^\d+$/;

            /**
             * Reset and set tab stops for a single paragraph.
             */
            function resetTabStops(paragraph, positions) {
                // Remove existing
                while (paragraph.tabStops.length > 0) {
                    paragraph.tabStops[0].remove();
                }
                // Add new
                for (var i = 0; i < positions.length; i++) {
                    paragraph.tabStops.add({
                        alignment: TabStopAlignment.LEFT_ALIGN,
                        position: positions[i]
                    });
                }
            }

            /**
             * Actually handle the text frame:
             *  1) Convert multiple spaces to tabs
             *  2) Apply tab stops
             *  3) Resize the frame
             */
            function handleNumbers(textFrame, newContent, config) {
                // 1) Normalizing content: multiple spaces => single tabs
                //    This can be adapted if you prefer no content changes.
                newContent = newContent.replace(/\s+/g, "\t");

                // 2) If the text frame contents differ, update it
                if (textFrame.contents !== newContent) {
                    textFrame.contents = newContent;
                    modifiedCount++;
                }

                // 3) Apply tab stops to each paragraph
                var paragraphs = textFrame.paragraphs;
                for (var j = 0; j < paragraphs.length; j++) {
                    paragraphs[j].justification = Justification.LEFT_ALIGN;
                    resetTabStops(paragraphs[j], config.tabPositions);
                }

                // 4) Resize frame width
                var gb = textFrame.geometricBounds; // [y1, x1, y2, x2]
                var top  = gb[0];
                var left = gb[1];
                var bottom = gb[2];
                var right = left + config.frameWidth;
                textFrame.geometricBounds = [top, left, bottom, right];

                return true;
            }

            // -------------------------------------------------
            // PROCESS EACH PAGE
            // -------------------------------------------------
            for (var pageIndex = 0; pageIndex < totalPages; pageIndex++) {
                var page = pages[pageIndex];
                var pageName = page.name;

                // ----- SKIP PAGES IF NEEDED -----
                if (skipThesePages.indexOf(pageName) >= 0) {
                    // Skip entirely
                    progressBar.value = pageIndex + 1;
                    progressText.text = "Skipping page " + pageName + "...";
                    progressWindow.update();
                    continue;
                }

                // Get the frames on this page
                var textFrames = page.textFrames;

                // Loop each text frame
                for (var tfIndex = 0; tfIndex < textFrames.length; tfIndex++) {
                    var textFrame = textFrames[tfIndex];

                    // If no content, skip
                    if (!textFrame.contents) continue;

                    // Trim leading/trailing whitespace
                    var content = textFrame.contents.replace(/^\s+|\s+$/g, "");

                    // Option 1: Identify frames by numeric patterns (like your original code).
                    //           If matched, we know how many tab stops to apply.
                    if (matchSix.test(content)) {
                        if (handleNumbers(textFrame, content, settings.sixNumbers)) updatedCount++;
                    } 
                    else if (matchFour.test(content)) {
                        if (handleNumbers(textFrame, content, settings.fourNumbers)) updatedCount++;
                    }
                    else if (matchThree.test(content)) {
                        if (handleNumbers(textFrame, content, settings.threeNumbers)) updatedCount++;
                    }
                    else if (matchTwo.test(content)) {
                        if (handleNumbers(textFrame, content, settings.twoNumbers)) updatedCount++;
                    }
                    else if (matchSingle.test(content)) {
                        if (handleNumbers(textFrame, content, settings.singleNumber)) updatedCount++;
                    }
                    else {
                        // Option 2: If you prefer to identify by frame width:
                        //   e.g., if ~74 mm => sixNumbers, if ~57 => fourNumbers, etc.
                        //   (You'd want to round or check a tolerance.)
                        /*
                        var gb = textFrame.geometricBounds;
                        var frameW = gb[3] - gb[1];
                        // Use approximate checks, e.g.:
                        if (Math.abs(frameW - 74) < 0.5) {
                            if (handleNumbers(textFrame, content, settings.sixNumbers)) updatedCount++;
                        }
                        else if (Math.abs(frameW - 57) < 0.5) {
                            if (handleNumbers(textFrame, content, settings.fourNumbers)) updatedCount++;
                        }
                        // etc...
                        */
                    }
                }

                // Update progress bar
                progressBar.value = pageIndex + 1;
                progressText.text = "Processing page " + pageName + " (" + (pageIndex+1) + "/" + totalPages + ")";
                progressWindow.update();
            }

            // -------------------------------------------------
            // WRAP-UP & RESTORE PREFERENCES
            // -------------------------------------------------
            app.scriptPreferences.enableRedraw = true;
            doc.viewPreferences.horizontalMeasurementUnits = originalHUnits;
            doc.viewPreferences.verticalMeasurementUnits = originalVUnits;
            doc.preflightOptions.preflightOff = false;

            progressWindow.close();

            // Final message
            if (updatedCount > 0) {
                alert(
                    updatedCount + " text frames were processed.\n" +
                    modifiedCount + " text frames were actually modified.\n" +
                    "The wand worked its magic!"
                );
            } else {
                alert("No text frames matching the criteria were found.");
            }

        } catch (e) {
            // In case of error
            alert("Oops, the magic wand encountered a problem: " + e.message);
        } finally {
            // Ensure the progress window closes even on error
            app.scriptPreferences.enableRedraw = true;
            if (progressWindow.visible) {
                progressWindow.close();
            }
        }
    },
    // UndoModes: If FAST_ENTIRE_SCRIPT doesn't work on your version,
    // use ENTIRE_SCRIPT instead.
    ScriptLanguage.JAVASCRIPT,
    [],
    UndoModes.FAST_ENTIRE_SCRIPT,
    "HandleNumbers-WithRangesAndWidths"
);
