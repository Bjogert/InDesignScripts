app.doScript(
    function() {
        // Turn off screen redraw for better performance
        app.scriptPreferences.enableRedraw = false;

        /*
        PlaceraTextrutorIPar-PerLager.jsx

        Justera dessa värden för att ändra avstånd och storlek:
        -------------------------------------------------------
        - pairSpacing: avstånd mellan två textrutor i samma par
        - groupSpacing: avstånd mellan paren
        - newHeight: höjd för varje textruta
        */

        if (app.selection.length < 1) {
            alert("Markera minst två textrutor för att använda detta skript.");
            // Exit early
            return;
        } else {
            // 1) Ange dina standardavstånd
            var pairSpacing = 0.963;  // Avstånd mellan två textrutor i samma par
            var groupSpacing = 5.94;  // Avstånd mellan par
            var newHeight   = 4;      // Sätt höjden på textrutorna till 4 mm

            // 2) Samla alla markerade TextFrames per lager
            var framesByLayer = {};
            var sel = app.selection;

            for (var i = 0; i < sel.length; i++) {
                if (sel[i].constructor.name === "TextFrame") {
                    var layerName = sel[i].itemLayer.name;
                    if (!framesByLayer[layerName]) {
                        framesByLayer[layerName] = [];
                    }
                    framesByLayer[layerName].push(sel[i]);
                }
            }

            // 3) För varje lager: sortera efter toppkant och använd "parvis placering"
            for (var layer in framesByLayer) {
                var textFrames = framesByLayer[layer];

                // Sortera efter top (y-koordinat)
                textFrames.sort(function(a, b) {
                    return a.geometricBounds[0] - b.geometricBounds[0];
                });

                // Placera textrutorna parvis
                for (var j = 0; j < textFrames.length; j++) {
                    var bounds = textFrames[j].geometricBounds;

                    // Ändra höjden utan att flytta sidled
                    var top = bounds[0];
                    var bottom = top + newHeight;
                    textFrames[j].geometricBounds = [top, bounds[1], bottom, bounds[3]];

                    // Justera vertikal placering (förutom första rutan i listan)
                    if (j > 0) {
                        var prevBounds = textFrames[j - 1].geometricBounds;
                        
                        if (j % 2 === 1) {
                            // Andra textrutan i ett par: flytta ner pairSpacing från föregående (första) ruta
                            var newTop = prevBounds[2] + pairSpacing;
                            var newBottom = newTop + newHeight;
                            textFrames[j].geometricBounds = [newTop, bounds[1], newBottom, bounds[3]];
                        } else {
                            // Första textrutan i nästa par: flytta ner groupSpacing från föregående par
                            var newTop = prevBounds[2] + groupSpacing;
                            var newBottom = newTop + newHeight;
                            textFrames[j].geometricBounds = [newTop, bounds[1], newBottom, bounds[3]];
                        }
                    }
                }
            }
            // Klart! Alla lager har nu fått parvis placering oberoende av varandra.
        }

        // Turn redraw back on now that we’re done
        app.scriptPreferences.enableRedraw = true;
    },
    ScriptLanguage.JAVASCRIPT,
    [],
    UndoModes.FAST_ENTIRE_SCRIPT, // Only one undo step for the entire script
    "PlaceraTextrutorIPar-PerLager - Single Undo"
);
