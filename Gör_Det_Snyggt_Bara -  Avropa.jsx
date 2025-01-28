// ===========================================================
// HandleNumbers-SkipPage74-Final-Working.jsx
// ===========================================================

app.doScript(
    function main() {
        var progressWindow = new Window("palette", "Sprinkling Fairy Dust", undefined, { closeButton: false });
        progressWindow.orientation = "column";
        progressWindow.alignChildren = ["fill", "center"];

        var progressBar = progressWindow.add("progressbar", undefined, 0, 100);
        progressBar.preferredSize.width = 300;
        var progressText = progressWindow.add("statictext", undefined, "Gathering magical energy...");
        progressText.alignment = ["fill", "center"];

        progressWindow.show();

        try {
            var doc = app.activeDocument;

            // ---------------------
            // Configurable Settings
            // ---------------------
            var specialPages = [40]; // index 41, det är sidan före som bestämmer. 
            var skipPages = [63]; // index 64, det är sidan före som bestämmer. 
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

            // Unlock and make all layers visible
            var layers = doc.layers;
            for (var i = 0; i < layers.length; i++) {
                layers[i].locked = false;
                layers[i].visible = true;
            }

            // Turn off Preflight for speed
            doc.preflightOptions.preflightOff = true;

            var originalHUnits = doc.viewPreferences.horizontalMeasurementUnits;
            var originalVUnits = doc.viewPreferences.verticalMeasurementUnits;
            doc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.MILLIMETERS;
            doc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.MILLIMETERS;

            app.scriptPreferences.enableRedraw = false;

            var pages = doc.pages;
            var totalPages = pages.length;
            var updatedCount = 0;
            var modifiedCount = 0;

            progressBar.maxvalue = totalPages;

            var matchSix = /^\d+\s+\d+\s+\d+\s+\d+\s+\d+\s+\d+$/;
            var matchFour = /^\d+\s+\d+\s+\d+\s+\d+$/;
            var matchThree = /^\d+\s+\d+\s+\d+$/;
            var matchTwo = /^\d+\s+\d+$/;
            var matchSingle = /^\d+$/;

            /**
             * Resets and sets tab stops for a paragraph based on provided positions.
             */
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

            /**
             * Handles the special case for three-number text frames on the special pages.
             */
            function handleSpecialCase(textFrame, config) {
                var bounds = textFrame.geometricBounds;
                var newLeft = bounds[1];
                var newRight = newLeft + config.frameWidth;
                textFrame.geometricBounds = [bounds[0], newLeft, bounds[2], newRight];

                var content = textFrame.contents.replace(/^\s+|\s+$/g, "");
                content = "\t" + content.replace(/\s+/g, "\t");

                if (textFrame.contents !== content) {
                    textFrame.contents = content;
                    modifiedCount++;
                }

                var paragraphs = textFrame.paragraphs;
                for (var j = 0; j < paragraphs.length; j++) {
                    paragraphs[j].justification = Justification.LEFT_ALIGN;
                    resetTabStops(paragraphs[j], config.tabPositions);
                }

                return true;
            }

            /**
             * Handles general cases for text frames with different number formats.
             */
            function handleNumbers(textFrame, content, config) {
                content = content.replace(/\s+/g, "\t");

                if (textFrame.contents !== content) {
                    textFrame.contents = content;
                    modifiedCount++;
                }

                var paragraphs = textFrame.paragraphs;
                for (var j = 0; j < paragraphs.length; j++) {
                    paragraphs[j].justification = Justification.LEFT_ALIGN;
                    resetTabStops(paragraphs[j], config.tabPositions);
                }

                var bounds = textFrame.geometricBounds;
                var newLeft = bounds[1];
                var newRight = newLeft + config.frameWidth;
                textFrame.geometricBounds = [bounds[0], newLeft, bounds[2], newRight];

                return true;
            }

            // Process pages one by one
            for (var pageIndex = 0; pageIndex < totalPages; pageIndex++) {
                // Skip this page if it's in the skipPages array
                var shouldSkip = false;
                for (var skipIndex = 0; skipIndex < skipPages.length; skipIndex++) {
                    if (skipPages[skipIndex] === pageIndex) {
                        shouldSkip = true;
                        break;
                    }
                }
                if (shouldSkip) {
                    progressText.text = "Skipping page " + (pageIndex + 1);
                    progressWindow.update();
                    continue;
                }

                var page = pages[pageIndex];
                var isSpecialPage = false;

                // Check if the current page is a special page
                for (var spIndex = 0; spIndex < specialPages.length; spIndex++) {
                    if (specialPages[spIndex] === pageIndex) {
                        isSpecialPage = true;
                        break;
                    }
                }

                var textFrames = page.textFrames;

                for (var tfIndex = 0; tfIndex < textFrames.length; tfIndex++) {
                    var textFrame = textFrames[tfIndex];

                    if (textFrame.contents) {
                        var content = textFrame.contents.replace(/^\s+|\s+$/g, "");
                        var bounds = textFrame.geometricBounds;
                        var frameWidth = bounds[3] - bounds[1];

                        if (isSpecialPage && matchThree.test(content) && Math.abs(frameWidth - 74) < 0.1) {
                            var changesMade = handleSpecialCase(textFrame, settings.sixNumbers);
                            if (changesMade) updatedCount++;
                        } else if (matchSix.test(content)) {
                            var changesMade = handleNumbers(textFrame, content, settings.sixNumbers);
                            if (changesMade) updatedCount++;
                        } else if (matchFour.test(content)) {
                            var changesMade = handleNumbers(textFrame, content, settings.fourNumbers);
                            if (changesMade) updatedCount++;
                        } else if (matchThree.test(content)) {
                            var changesMade = handleNumbers(textFrame, content, settings.threeNumbers);
                            if (changesMade) updatedCount++;
                        } else if (matchTwo.test(content)) {
                            var changesMade = handleNumbers(textFrame, content, settings.twoNumbers);
                            if (changesMade) updatedCount++;
                        } else if (matchSingle.test(content)) {
                            var changesMade = handleNumbers(textFrame, content, settings.singleNumber);
                            if (changesMade) updatedCount++;
                        }
                    }
                }

                progressBar.value = pageIndex + 1;
                progressText.text = "Processing page " + (pageIndex + 1) + " of " + totalPages;
                progressWindow.update();
            }

            app.scriptPreferences.enableRedraw = true;

            if (updatedCount > 0) {
                alert(updatedCount + " text frames were processed.\n" +
                      modifiedCount + " text frames were actually modified. The wand worked its magic!");
            } else {
                alert("No text frames matching the criteria were found.");
            }

            doc.viewPreferences.horizontalMeasurementUnits = originalHUnits;
            doc.viewPreferences.verticalMeasurementUnits = originalVUnits;
            doc.preflightOptions.preflightOff = false;

            progressWindow.close();

        } catch (e) {
            alert("Oops, the magic wand encountered a problem: " + e.message);
        } finally {
            app.scriptPreferences.enableRedraw = true;

            if (progressWindow.visible) {
                progressWindow.close();
            }
        }
    },
    ScriptLanguage.JAVASCRIPT,
    [],
    UndoModes.FAST_ENTIRE_SCRIPT,
    "Handle Numbers with Special Pages and Skipping"
);
