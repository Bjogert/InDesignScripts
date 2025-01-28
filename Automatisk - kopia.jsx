// ===========================================================
// HandleNumbers-WithSpecialPages-Final-Simplified.jsx
// ===========================================================

app.doScript(
    function main() {
        var progressWindow = new Window("palette", "Processing Text Frames", undefined, { closeButton: false });
        progressWindow.orientation = "column";
        progressWindow.alignChildren = ["fill", "center"];

        var progressBar = progressWindow.add("progressbar", undefined, 0, 100);
        progressBar.preferredSize.width = 300;
        var progressText = progressWindow.add("statictext", undefined, "Initializing...");
        progressText.alignment = ["fill", "center"];

        progressWindow.show();

        try {
            var doc = app.activeDocument;

            // ---------------------
            // Configurable Settings
            // ---------------------
            var specialPages = [48]; // Pages 49 (index 48)
            var settings = {
                sixNumbers: {
                    tabPositions: [13.05, 26.1, 39.1, 51.86, 64.68], // Tab positions in mm
                    frameWidth: 74                                // Frame width in mm
                },
                fourNumbers: {
                    tabPositions: [15, 30, 45],                  // Tab positions in mm
                    frameWidth: 57                                // Frame width in mm
                },
                threeNumbers: {
                    tabPositions: [15, 30, 45],                  // Tab positions in mm
                    frameWidth: 42                                // Frame width in mm
                },
                twoNumbers: {
                    tabPositions: [26],                          // Tab positions in mm
                    frameWidth: 35                                // Frame width in mm
                },
                singleNumber: {
                    tabPositions: [0],                           // Tab position in mm
                    frameWidth: 10                                // Frame width in mm
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

            var matchSix = /^\d+\s+\d+\s+\d+\s+\d+\s+\d+\s+\d+$/; // Match exactly six digits separated by spaces
            var matchFour = /^\d+\s+\d+\s+\d+\s+\d+$/;            // Match exactly four digits separated by spaces
            var matchThree = /^\d+\s+\d+\s+\d+$/;                // Match exactly three digits separated by spaces
            var matchTwo = /^\d+\s+\d+$/;                        // Match exactly two digits separated by spaces
            var matchSingle = /^\d+$/;                           // Match exactly one digit

            /**
             * Resets and sets tab stops for a paragraph based on provided positions.
             * @param {Paragraph} paragraph - The paragraph to modify.
             * @param {Array} positions - An array of numerical positions (in mm) for tab stops.
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
             * @param {TextFrame} textFrame - The text frame to modify.
             * @param {Object} config - The tab positions and frame width for this case.
             */
            function handleSpecialCase(textFrame, config) {
                // Format the box width to 74mm
                var bounds = textFrame.geometricBounds;
                var newLeft = bounds[1];
                var newRight = newLeft + config.frameWidth; // Set exact width to 74mm
                textFrame.geometricBounds = [bounds[0], newLeft, bounds[2], newRight];

                // Remove all old tabs and add one tab before each number
                var content = textFrame.contents.replace(/^\s+|\s+$/g, ""); // Trim whitespace
                content = "\t" + content.replace(/\s+/g, "\t"); // Add one tab before each number

                if (textFrame.contents !== content) {
                    textFrame.contents = content; // Update the content
                    modifiedCount++; // Increment modified count if content changes
                }

                // Set tab stops and justification
                var paragraphs = textFrame.paragraphs;
                for (var j = 0; j < paragraphs.length; j++) {
                    paragraphs[j].justification = Justification.LEFT_ALIGN;
                    resetTabStops(paragraphs[j], config.tabPositions); // Use specified tab positions
                }

                return true; // Changes made
            }

            /**
             * Handles general cases for text frames with different number formats.
             * @param {TextFrame} textFrame - The text frame to modify.
             * @param {String} content - The sanitized text content.
             * @param {Object} config - The tab positions and frame width for this case.
             */
            function handleNumbers(textFrame, content, config) {
                content = content.replace(/\s+/g, "\t"); // Replace all whitespaces with tabs

                if (textFrame.contents !== content) {
                    textFrame.contents = content;
                    modifiedCount++; // Increment modified count if content changes
                }

                // Set tab stops and justification
                var paragraphs = textFrame.paragraphs;
                for (var j = 0; j < paragraphs.length; j++) {
                    paragraphs[j].justification = Justification.LEFT_ALIGN;
                    resetTabStops(paragraphs[j], config.tabPositions);
                }

                // Adjust geometric bounds to the specified frame width
                var bounds = textFrame.geometricBounds;
                var newLeft = bounds[1];
                var newRight = newLeft + config.frameWidth;
                textFrame.geometricBounds = [bounds[0], newLeft, bounds[2], newRight];

                return true; // Changes made
            }

            // Process pages one by one
            for (var pageIndex = 0; pageIndex < totalPages; pageIndex++) {
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
                        var content = textFrame.contents.replace(/^\s+|\s+$/g, ""); // Trim whitespace
                        var bounds = textFrame.geometricBounds;
                        var frameWidth = bounds[3] - bounds[1]; // Width = right - left

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

                // Update progress bar after each page
                progressBar.value = pageIndex + 1;
                progressText.text = "Processing page " + (pageIndex + 1) + " of " + totalPages;
                progressWindow.update();
            }

            app.scriptPreferences.enableRedraw = true;

            // Alert user about results
            if (updatedCount > 0) {
                alert(updatedCount + " text frames were processed.\n" +
                      modifiedCount + " text frames were actually modified.");
            } else {
                alert("No text frames matching the criteria were found.");
            }

            doc.viewPreferences.horizontalMeasurementUnits = originalHUnits;
            doc.viewPreferences.verticalMeasurementUnits = originalVUnits;
            doc.preflightOptions.preflightOff = false;

            progressWindow.close();

        } catch (e) {
            alert("An error occurred: " + e.message);
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
    "Handle Numbers with Special Pages"
);
