// FlyttaTextTillVänsterOchTextrutanHöger.jsx
// -------------

// Kontrollera om något är markerat
if (app.selection.length > 0) {
    var updatedCount = 0;

    for (var i = 0; i < app.selection.length; i++) {
        var textFrame = app.selection[i];

        // Kontrollera om objektet är en textruta och innehåller text
        if (textFrame.constructor.name === "TextFrame" && textFrame.contents) {
            // Sätt textjusteringen till vänsterjusterad
            var paragraphs = textFrame.paragraphs;
            for (var j = 0; j < paragraphs.length; j++) {
                paragraphs[j].justification = Justification.LEFT_ALIGN;
            }

            // Ta bort eventuella tabbar eller mellanslag i början av texten
            textFrame.contents = textFrame.contents.replace(/^\s+/, '');

            // Flytta textrutan 5 mm till höger
            var bounds = textFrame.geometricBounds; // Hämta nuvarande gränser
            var top = bounds[0];
            var left = bounds[1] + 5.2; // Flytta vänsterkanten 5 mm
            var bottom = bounds[2];
            var right = bounds[3] + 5.2; // Flytta högerkanten 5 mm
            textFrame.geometricBounds = [top, left, bottom, right];

            updatedCount++;
        }
    }

    // Visa meddelande om hur många textrutor som uppdaterades
    if (updatedCount > 0) {
        alert(updatedCount + " textrutor flyttades till vänster och textrutorna 5 mm till höger.");
    } else {
        alert("Inga textrutor uppdaterades.");
    }
} else {
    alert("Markera en eller flera textrutor innan du kör skriptet.");
}
