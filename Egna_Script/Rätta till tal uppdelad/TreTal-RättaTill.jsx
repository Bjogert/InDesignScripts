// TabbaTreTal-MedTabbstopp-Och38mm.jsx
// -------------

// Nollställ eventuella gamla inställningar för GREP
app.findGrepPreferences = NothingEnum.nothing;
app.changeGrepPreferences = NothingEnum.nothing;

// Sök efter exakt tre tal separerade av mellanslag
app.findGrepPreferences.findWhat = '^(\\d+)\\s+(\\d+)\\s+(\\d+)$';

if (app.selection.length > 0) {
    var updatedCount = 0;

    for (var i = 0; i < app.selection.length; i++) {
        var textFrame = app.selection[i];

        // Kontrollera om objektet är en textruta och innehåller text
        if (textFrame.constructor.name === "TextFrame" && textFrame.contents) {
            // Sök efter matchning i textrutan
            var matches = textFrame.findGrep();
            if (matches.length > 0) {
                for (var j = 0; j < matches.length; j++) {
                    // Utför GREP-ersättning för att lägga till tabbar
                    app.changeGrepPreferences.changeTo = '$1\\t$2\\t$3';
                    matches[j].changeGrep();

                    // Ställ in specifika tabbstopp
                    var paragraph = matches[j].paragraphs[0];
                    paragraph.tabStops.everyItem().position = 0; // Nollställ tidigare tabbstopp

                    // Lägg till nya tabbstopp
                    paragraph.tabStops.add({
                        alignment: TabStopAlignment.LEFT_ALIGN,
                        position: 15 // Första tabbens avstånd i punkter
                    });
                    paragraph.tabStops.add({
                        alignment: TabStopAlignment.LEFT_ALIGN,
                        position: 26.5 // Andra tabbens avstånd i punkter
                    });

                    // Ställ in styckejusteringen till "Marginaljustera med sista raden centrerad"
                    paragraph.justification = Justification.CENTER_JUSTIFIED;

                    // Sätt textrutans bredd till exakt 38 mm
                    var bounds = textFrame.geometricBounds; // Hämta nuvarande geometriska gränser
                    var top = bounds[0]; // Övre kant
                    var left = bounds[1]; // Vänster kant
                    var bottom = bounds[2]; // Nedre kant
                    var right = left + 38; // Höger kant (vänster + 38 mm)

                    textFrame.geometricBounds = [top, left, bottom, right]; // Tilldela nya gränser

                    updatedCount++;
                }
            }
        }
    }

    if (updatedCount > 0) {
        alert(updatedCount + " textrutor uppdaterades med tabbstopp, justering och bredd.");
    } else {
        alert("Inga textrutor med exakt tre tal hittades i de markerade objekten.");
    }
} else {
    alert("Markera en eller flera textrutor innan du kör skriptet.");
}

// Återställ GREP-inställningar
app.findGrepPreferences = NothingEnum.nothing;
app.changeGrepPreferences = NothingEnum.nothing;


