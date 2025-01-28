// @ts-nocheck #target "InDesign"

/*
    MakeItNeat-MainScript.jsx

    This script processes an InDesign document in one undo step. It:

    1) Unlocks and makes all layers visible.
    2) Turns off preflight for speed.
    3) Loops through every page in the document.
    4) For pages listed in "leftAlignPages", applies a
       "LeftAlignNumbersWithFrameWidth" routine:
       - Only numeric frames (digits + whitespace) are processed.
       - Old tabs/spaces are cleaned up.
       - The text is split into chunks, then rejoined with single tabs.
       - Left-aligned tab stops are added, and the frame is resized on the right side.
    5) For other pages, it uses your original logic:
       - Optionally checks "specialPages" for a custom handleSpecialCase.
       - Otherwise applies handleNumbers to frames matching certain regex patterns
         (six, four, three, two, single digits).
    6) Shows a progress bar and processes everything in ONE undo step.

    Customize:
      - specialPages: array of page indices with special-case logic.
      - leftAlignPages: array of page indices that will use the chunk-based left-align routine.
      - Edit "settings" for your original handleNumbers approach.
      - Edit "leftAlignSettingsByChunkCount" for the chunk-based logic's tab positions/frame widths.
      - Indices are zero-based (page 1 => index 0, page 2 => index 1, etc.).
*/

app.doScript(
    function main() {
        // ------------------------------------------------
        // 0) CONFIG
        // ------------------------------------------------

        // Page indices needing a special approach
        // (uses handleSpecialCase if certain conditions match)
        var specialPages = [50]; // Example: page 49 (index 48)

        // The page(s) that should use the "LeftAlignNumbersWithFrameWidth" approach
        // Example: page 74 => index 73
        var leftAlignPages = [75];

        // Original "handleNumbers" settings (the old approach).
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

        // LeftAlignNumbersWithFrameWidth approach (chunk-based logic)
        var leftAlignSettingsByChunkCount = {
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
            // Extend if needed (5, 6, etc.)
        };

        // ------------------------------------------------
        // 1) INITIAL SETUP
        // ------------------------------------------------

        if (app.documents.length === 0) {
            alert("No document open!");
            return;
        }
        var doc = app.activeDocument;

        // Unlock and make all layers visible
        var layers = doc.layers;
        for (var l = 0; l < layers.length; l++) {
            layers[l].locked = false;
            layers[l].visible = true;
        }

        // Turn off preflight for performance
        doc.preflightOptions.preflightOff = true;

        // Store old units and switch to millimeters
        var oldH = doc.viewPreferences.horizontalMeasurementUnits;
        var oldV = doc.viewPreferences.verticalMeasurementUnits;
        doc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.MILLIMETERS;
        doc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.MILLIMETERS;

        // Turn off redraw for speed
        app.scriptPreferences.enableRedraw = false;

        // ------------------------------------------------
        // 2) CREATE PROGRESS UI
        // ------------------------------------------------
        var w = new Window("palette", "Make it neat and tidy", undefined, { closeButton: false });
        w.orientation = "column";
        w.alignChildren = ["fill", "center"];

        var progressBar = w.add("progressbar", undefined, 0, 100);
        progressBar.preferredSize.width = 300;
        var progressText = w.add("statictext", undefined, "Gathering magical energy...");
        progressText.alignment = ["fill", "center"];

        w.show();

        // ------------------------------------------------
        // 3) PAGE LOOP
        // ------------------------------------------------
        var pages = doc.pages;
        var totalPages = pages.length;
        progressBar.maxvalue = totalPages;

        var updatedCount = 0;
        var modifiedCount = 0;

        // Regex patterns for the original approach
        var matchSix =   /^\d+\s+\d+\s+\d+\s+\d+\s+\d+\s+\d+$/;
        var matchFour =  /^\d+\s+\d+\s+\d+\s+\d+$/;
        var matchThree = /^\d+\s+\d+\s+\d+$/;
        var matchTwo =   /^\d+\s+\d+$/;
        var matchSingle =/^\d+$/;

        // ------------------------------------------------
        // 4) HELPER FUNCTIONS
        // ------------------------------------------------

        // Original "handleNumbers" function
        function handleNumbers(textFrame, content, config) {
            // Convert spaces to tabs
            content = content.replace(/\s+/g, "\t");

            if (textFrame.contents !== content) {
                textFrame.contents = content;
                modifiedCount++;
            }

            // Left-align paragraphs, apply tab stops
            var paragraphs = textFrame.paragraphs;
            for (var j = 0; j < paragraphs.length; j++) {
                paragraphs[j].justification = Justification.LEFT_ALIGN;
                resetTabStops(paragraphs[j], config.tabPositions);
            }

            // Resize frame: keep left edge, extend right
            var bounds = textFrame.geometricBounds; // [top, left, bottom, right]
            var newLeft = bounds[1];
            var newRight = newLeft + config.frameWidth;
            textFrame.geometricBounds = [bounds[0], newLeft, bounds[2], newRight];

            return true;
        }

        // Original "handleSpecialCase" function
        function handleSpecialCase(textFrame, config) {
            var bounds = textFrame.geometricBounds;
            var newLeft = bounds[1];
            var newRight = newLeft + config.frameWidth;
            textFrame.geometricBounds = [bounds[0], newLeft, bounds[2], newRight];

            var content = textFrame.contents.replace(/^\s+|\s+$/g, "");
            // Insert a leading tab, then replace spaces with tabs
            content = "\t" + content.replace(/\s+/g, "\t");

            if (textFrame.contents !== content) {
                textFrame.contents = content;
                modifiedCount++;
            }

            // Left-align paragraphs, apply tab stops
            var paragraphs = textFrame.paragraphs;
            for (var j = 0; j < paragraphs.length; j++) {
                paragraphs[j].justification = Justification.LEFT_ALIGN;
                resetTabStops(paragraphs[j], config.tabPositions);
            }

            return true;
        }

        // Reset tab stops helper
        function resetTabStops(paragraph, positions) {
            while (paragraph.tabStops.length > 0) {
                paragraph.tabStops[0].remove();
            }
            for (var i = 0; i < positions.length; i++) {
                paragraph.tabStops.add({
                    alignment: TabStopAlignment.LEFT_ALIGN,
                    position: positions[i]
                });
            }
        }

        // "LeftAlignNumbersWithFrameWidth" for a single page
        function runLeftAlignNumbersOnPage(doc, pageIndex, configByCount) {
            var thePage = doc.pages[pageIndex];
            var frames = thePage.textFrames;
            for (var f = 0; f < frames.length; f++) {
                var tf = frames[f];
                var didChange = processChunkedFrame(tf, configByCount);
                if (didChange) {
                    updatedCount++;
                    modifiedCount++;
                }
            }
        }

        /*
            processChunkedFrame:
            1) Checks if text is purely numeric (digits + whitespace).
            2) Cleans up tabs/spaces.
            3) Splits text into chunks, re-inserts single tabs.
            4) Applies left-aligned tab stops & resizes frame (extend to right).
        */
        function processChunkedFrame(textFrame, settingsByChunkCount) {
            if (!textFrame.isValid || !textFrame.parentStory) return false;

            var oldContent = textFrame.contents;
            if (!oldContent) return false;

            // Remove existing tabs, collapse multiple spaces
            var newContent = oldContent.replace(/\t+/g, " ");
            newContent = newContent.replace(/\s+/g, " ");
            newContent = newContent.replace(/^\s+|\s+$/g, "");

            // Must be digits + spaces only
            var isNumericOnly = /^[0-9 ]*$/.test(newContent);
            if (!isNumericOnly) return false;
            if (!newContent) return false;

            // Split => chunks
            var chunks = newContent.split(" ");
            // Rejoin with single tabs
            newContent = chunks.join("\t");

            var contentChanged = (newContent !== oldContent);
            if (contentChanged) {
                textFrame.contents = newContent;
            }

            // Choose config based on chunk count
            var chunkCount = chunks.length;
            var config = settingsByChunkCount[chunkCount];
            if (!config) {
                return contentChanged; // no config => do nothing more
            }

            // Apply left-aligned tab stops
            var paragraphs = textFrame.parentStory.paragraphs;
            for (var p = 0; p < paragraphs.length; p++) {
                var para = paragraphs[p];
                while (para.tabStops.length > 0) {
                    para.tabStops[0].remove();
                }
                var positions = config.tabPositions || [];
                for (var t = 0; t < positions.length; t++) {
                    var ts = para.tabStops.add();
                    ts.alignment = TabStopAlignment.LEFT_ALIGN;
                    ts.position = positions[t];
                }
                para.justification = Justification.LEFT_ALIGN;
            }

            // Resize frame by extending right side
            if (config.frameWidth != null) {
                var gb = textFrame.geometricBounds; // [top, left, bottom, right]
                var top    = gb[0];
                var left   = gb[1];
                var bottom = gb[2];
                var right  = left + config.frameWidth;
                textFrame.geometricBounds = [top, left, bottom, right];
            }

            return contentChanged;
        }

        // ------------------------------------------------
        // 5) PROCESS PAGES
        // ------------------------------------------------
        for (var pageIndex = 0; pageIndex < totalPages; pageIndex++) {
            // Check if this page is in leftAlignPages using a for-loop
            var isLeftAlignPage = false;
            for (var la = 0; la < leftAlignPages.length; la++) {
                if (leftAlignPages[la] === pageIndex) {
                    isLeftAlignPage = true;
                    break;
                }
            }

            if (isLeftAlignPage) {
                // If yes, run the chunk-based logic on that page only
                runLeftAlignNumbersOnPage(doc, pageIndex, leftAlignSettingsByChunkCount);
                progressBar.value = pageIndex + 1;
                progressText.text = "Left-align logic on page " + (pageIndex + 1);
                w.update();
                continue;
            }

            // Otherwise, do the original approach
            var page = pages[pageIndex];

            // Check if page is special
            var isSpecialPage = false;
            for (var spIndex = 0; spIndex < specialPages.length; spIndex++) {
                if (specialPages[spIndex] === pageIndex) {
                    isSpecialPage = true;
                    break;
                }
            }

            var textFrames = page.textFrames;
            for (var tfIndex = 0; tfIndex < textFrames.length; tfIndex++) {
                var textFrame = textFrames[tfIndex];
                if (!textFrame.contents) continue;

                var content = textFrame.contents.replace(/^\s+|\s+$/g, "");
                var gb = textFrame.geometricBounds; // [top, left, bottom, right]
                var frameWidth = gb[3] - gb[1];

                // Original logic checks
                if (isSpecialPage && matchThree.test(content) && Math.abs(frameWidth - 74) < 0.1) {
                    var changedA = handleSpecialCase(textFrame, settings.sixNumbers);
                    if (changedA) updatedCount++;
                } else if (matchSix.test(content)) {
                    var changedB = handleNumbers(textFrame, content, settings.sixNumbers);
                    if (changedB) updatedCount++;
                } else if (matchFour.test(content)) {
                    var changedC = handleNumbers(textFrame, content, settings.fourNumbers);
                    if (changedC) updatedCount++;
                } else if (matchThree.test(content)) {
                    var changedD = handleNumbers(textFrame, content, settings.threeNumbers);
                    if (changedD) updatedCount++;
                } else if (matchTwo.test(content)) {
                    var changedE = handleNumbers(textFrame, content, settings.twoNumbers);
                    if (changedE) updatedCount++;
                } else if (matchSingle.test(content)) {
                    var changedF = handleNumbers(textFrame, content, settings.singleNumber);
                    if (changedF) updatedCount++;
                }
            }

            // Update progress
            progressBar.value = pageIndex + 1;
            progressText.text = "Sprinkling fairydust on page " + (pageIndex + 1) + " of " + totalPages;
            w.update();
        }

        // ------------------------------------------------
        // 6) WRAP UP
        // ------------------------------------------------
        app.scriptPreferences.enableRedraw = true;
        doc.preflightOptions.preflightOff = false;
        doc.viewPreferences.horizontalMeasurementUnits = oldH;
        doc.viewPreferences.verticalMeasurementUnits = oldV;

        w.close();

        if (updatedCount > 0) {
            alert(
                updatedCount + " text frames were processed.\n" +
                modifiedCount + " text frames were actually tinkered with.\n" +
                "Am i a good computer or just OK?"
            );
        } else {
            alert("No text frames matching the criteria were found.");
        }
    },
    ScriptLanguage.JAVASCRIPT,
    [],
    // Single undo step
    UndoModes.FAST_ENTIRE_SCRIPT,
    "MakeItNeat-MainScript"
);
