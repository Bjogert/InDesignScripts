// ===========================================================
// RestoreSelectedTextFrames-AdjustableTabDistance.jsx
// ===========================================================

app.doScript(
    function main() {
        try {
            var doc = app.activeDocument;

            if (!doc.selection || doc.selection.length === 0) {
                alert("Please select one or more text frames to restore.");
                return;
            }

            var selection = doc.selection;
            var restoredCount = 0;

            // ---------------------
            // Configurable Settings
            // ---------------------
            var settings = {
                tabPositions: [20, 40], // Tab distances in mm (adjustable)
                frameWidth: 30          // Default frame width for two-number boxes
            };

            /**
             * Resets and sets tab stops for a paragraph.
             * @param {Paragraph} paragraph - The paragraph to modify.
             * @param {Array} positions - An array of numerical positions for tab stops.
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
             * Restores the selected text frame to its original state with adjustable tab distance.
             * @param {TextFrame} textFrame - The text frame to restore.
             * @param {Object} config - Configuration with tab stops and width.
             */
            function restoreTextFrame(textFrame, config) {
                // Remove all tabs and replace with spaces
                var content = textFrame.contents.replace(/\t+/g, " "); // Replace tabs with spaces

                if (textFrame.contents !== content) {
                    textFrame.contents = content; // Update content
                }

                // Reset tab stops and justification
                var paragraphs = textFrame.paragraphs;
                for (var j = 0; j < paragraphs.length; j++) {
                    paragraphs[j].justification = Justification.LEFT_ALIGN;
                    resetTabStops(paragraphs[j], config.tabPositions); // Use custom tab distances
                }

                // Reset geometric bounds to original width
                var bounds = textFrame.geometricBounds;
                var newLeft = bounds[1];
                var newRight = newLeft + config.frameWidth;
                textFrame.geometricBounds = [bounds[0], newLeft, bounds[2], newRight];

                restoredCount++;
            }

            // Process selected items
            for (var i = 0; i < selection.length; i++) {
                var selectedItem = selection[i];

                // Check if the selected item is a text frame
                if (selectedItem.constructor.name === "TextFrame") {
                    restoreTextFrame(selectedItem, settings);
                }
            }

            // Alert user about results
            if (restoredCount > 0) {
                alert(
                    restoredCount +
                        " selected text frames were restored to their original state with updated tab distances."
                );
            } else {
                alert("No selected text frames required restoration.");
            }
        } catch (e) {
            alert("Oops, something went wrong during restoration: " + e.message);
        }
    },
    ScriptLanguage.JAVASCRIPT,
    [],
    UndoModes.FAST_ENTIRE_SCRIPT,
    "Restore Selected Text Frames with Adjustable Tabs"
);
