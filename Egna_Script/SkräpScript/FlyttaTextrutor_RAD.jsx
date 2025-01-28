// JusteraTextrutor-MedHöjdOchPlacering-UtanMeddelande.jsx
// --------------------------------------------------------------------

//-0.585;0.12

// Justera dessa värden för att ändra standardhöjd och avstånd:
var newHeight = 4; // Textrutans höjd i mm
var distanceWithinLayer = 7.05; // Avståndet mellan textrutor i samma lager i mm

// Starta en ”ångra”-kontroll
app.doScript(function () {
    // Kontrollera om något är markerat
    if (app.selection.length < 1) {
        alert("Markera minst en textruta för att använda detta skript.");
        return;
    }

    // 1) Gruppera alla markerade textrutor efter lager
    var framesByLayer = {};
    var sel = app.selection;

    // Samla TextFrames lager för lager
    for (var i = 0; i < sel.length; i++) {
        if (sel[i].constructor.name === "TextFrame") {
            var layerName = sel[i].itemLayer.name;
            if (!framesByLayer[layerName]) {
                framesByLayer[layerName] = [];
            }
            framesByLayer[layerName].push(sel[i]);
        }
    }

    // 2) Hantera varje lager separat
    for (var layer in framesByLayer) {
        var textFrames = framesByLayer[layer];

        // Sortera textrutorna efter toppkant (y-koordinat)
        textFrames.sort(function (a, b) {
            return a.geometricBounds[0] - b.geometricBounds[0];
        });

        // Justera höjd och placering av textrutorna
        for (var j = 0; j < textFrames.length; j++) {
            var bounds = textFrames[j].geometricBounds;

            // Sätt textrutans höjd till newHeight
            var top = bounds[0];
            var bottom = top + newHeight;
            textFrames[j].geometricBounds = [top, bounds[1], bottom, bounds[3]];

            // Flytta ned textrutan om det inte är den första i lagret
            if (j > 0) {
                var prevBounds = textFrames[j - 1].geometricBounds;
                var newTop = prevBounds[2] + distanceWithinLayer; // Nederkant + avstånd
                var newBottom = newTop + newHeight;
                textFrames[j].geometricBounds = [newTop, bounds[1], newBottom, bounds[3]];
            }
        }
    }
}, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT);
