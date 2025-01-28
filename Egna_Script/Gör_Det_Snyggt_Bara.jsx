// AutomatisktHanteraTextrutor-ExaktLösningPerFall.jsx
// -------------

app.doScript(
    function() {
        // 1. Disable screen redraw at the very start
        app.scriptPreferences.enableRedraw = false;

// Kontrollera om något är markerat
if (app.selection.length > 0) {
    var updatedCount = 0;

    for (var i = 0; i < app.selection.length; i++) {
        var textFrame = app.selection[i];

        // Kontrollera om objektet är en textruta och innehåller text
        if (textFrame.constructor.name === "TextFrame" && textFrame.contents) {
            // Hämta texten i textrutan och trimma extra mellanslag
            var content = textFrame.contents.replace(/^\s+|\s+$/g, '');

            // Kontrollera antalet tal i textrutan
            var matchSingle = /^\d+$/; // Ett ensamt tal
            var matchTwo = /^\d+\s+\d+$/;
            var matchThree = /^\d+\s+\d+\s+\d+$/;
            var matchFour = /^\d+\s+\d+\s+\d+\s+\d+$/;
            var matchSix = /^\d+\s+\d+\s+\d+\s+\d+\s+\d+\s+\d+$/;

            if (matchSingle.test(content)) {
                // Hantera ett ensamt tal
                handleSingleNumber(textFrame);
                updatedCount++;
            } else if (matchTwo.test(content)) {
                // Hantera två tal
                handleTwoNumbers(textFrame);
                updatedCount++;
            } else if (matchThree.test(content)) {
                // Hantera tre tal
                handleThreeNumbers(textFrame);
                updatedCount++;
            } else if (matchFour.test(content)) {
                // Hantera fyra tal
                handleFourNumbers(textFrame);
                updatedCount++;
            } else if (matchSix.test(content)) {
                // Hantera sex tal
                handleSixNumbers(textFrame);
                updatedCount++;
            }
        }
    }

    // Visa meddelande om hur många textrutor som uppdaterades
    if (updatedCount > 0) {
        alert(updatedCount + " textrutor uppdaterades baserat på antal tal.");
    } else {
        alert("Inga textrutor med giltigt innehåll hittades.");
    }
} else {
    alert("Markera en eller flera textrutor innan du kör skriptet.");
}

// Hantering för textrutor med ett ensamt tal
function handleSingleNumber(textFrame) {
    // Kontrollera om texten redan är vänsterjusterad
    var paragraphs = textFrame.paragraphs;
    var isLeftAligned = true;
    for (var j = 0; j < paragraphs.length; j++) {
        if (paragraphs[j].justification !== Justification.LEFT_ALIGN) {
            isLeftAligned = false;
            break;
        }
    }

    // Om texten redan är vänsterjusterad, avbryt ytterligare justering
    if (isLeftAligned) {
        return; // Hoppa över flytt och storleksändring
    }

    // Hämta ursprungliga gränser
    var bounds = textFrame.geometricBounds;
    var top = bounds[0];
    var left = bounds[1];
    var bottom = bounds[2];
    var right = bounds[3];

    // Beräkna mittpunkten för textrutan
    var centerX = (left + right) / 2;

    // Justera textrutans bredd till exakt 15 mm med texten kvar på samma position
    var newLeft = centerX - 7.5; // Flytta vänsterkanten
    var newRight = centerX + 7.5; // Flytta högerkanten
    textFrame.geometricBounds = [top, newLeft, bottom, newRight];

    // Sätt textjusteringen till vänsterjusterad
    for (var j = 0; j < paragraphs.length; j++) {
        paragraphs[j].justification = Justification.LEFT_ALIGN;
    }

    // Flytta hela textrutan 5.2 mm till höger
    var finalLeft = newLeft + 5.2; // Flytta vänsterkanten
    var finalRight = newRight + 5.2; // Flytta högerkanten
    textFrame.geometricBounds = [top, finalLeft, bottom, finalRight];
}

// Hantering för två tal
function handleTwoNumbers(textFrame) {
    // Lägg till tre tabbar mellan talen
    textFrame.contents = textFrame.contents.replace(/\s+/, '\t\t\t');

    // Ändra justering till "Marginaljustera med sista raden vänsterjusterad"
    var paragraphs = textFrame.paragraphs;
    for (var j = 0; j < paragraphs.length; j++) {
        paragraphs[j].justification = Justification.LEFT_JUSTIFIED;
    }

    // Justera textrutans bredd till exakt 35 mm
    var bounds = textFrame.geometricBounds;
    bounds[3] = bounds[1] + 35; // Sätt högerkanten till vänsterkanten + 35 mm
    textFrame.geometricBounds = bounds;
}

// Hantering för tre tal
function handleThreeNumbers(textFrame) {
    // Nollställ eventuella gamla inställningar för GREP
    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;

    // Sök efter exakt tre tal separerade av mellanslag
    app.findGrepPreferences.findWhat = '^(\\d+)\\s+(\\d+)\\s+(\\d+)$';

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
                position: 27.8 // Andra tabbens avstånd i punkter
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
        }
    }

    // Återställ GREP-inställningar
    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;
}



// Hantering för fyra tal
function handleFourNumbers(textFrame) {
    // Nollställ eventuella gamla inställningar för GREP
    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;

    // Sök efter exakt fyra tal separerade av mellanslag
    app.findGrepPreferences.findWhat = '^(\\d+)\\s+(\\d+)\\s+(\\d+)\\s+(\\d+)$';

    // Sök efter matchning i textrutan
    var matches = textFrame.findGrep();
    if (matches.length > 0) {
        for (var j = 0; j < matches.length; j++) {
            // Utför GREP-ersättning för att lägga till tabbar
            app.changeGrepPreferences.changeTo = '$1\\t$2\\t$3\\t$4';
            matches[j].changeGrep();

            // Kontrollera om tabbstoppen redan är inställda
            var paragraph = matches[j].paragraphs[0];
            if (paragraph.tabStops.length > 0 &&
                paragraph.tabStops[0].position === 15 &&
                paragraph.tabStops[1].position === 30 &&
                paragraph.tabStops[2].position === 40) {
                continue; // Hoppa över om tabbstoppen redan är korrekt inställda
            }

            // Nollställ tidigare tabbstopp
            paragraph.tabStops.everyItem().position = 0;

            // Lägg till nya tabbstopp
            paragraph.tabStops.add({
                alignment: TabStopAlignment.LEFT_ALIGN,
                position: 15 // Första tabbens avstånd i punkter
            });
            paragraph.tabStops.add({
                alignment: TabStopAlignment.LEFT_ALIGN,
                position: 30 // Andra tabbens avstånd i punkter
            });
            paragraph.tabStops.add({
                alignment: TabStopAlignment.LEFT_ALIGN,
                position: 40 // Tredje tabbens avstånd i punkter
            });

            // Ställ in styckejusteringen till "Marginaljustera med sista raden centrerad"
            paragraph.justification = Justification.CENTER_JUSTIFIED;

            // Kontrollera om textrutan redan är 55 mm bred
            var bounds = textFrame.geometricBounds; // Hämta nuvarande gränser
            var currentWidth = bounds[3] - bounds[1]; // Bredd på textrutan
            if (currentWidth !== 55) {
                // Justera bredden om den inte redan är korrekt
                var top = bounds[0]; // Övre kant
                var left = bounds[1]; // Vänster kant
                var bottom = bounds[2]; // Nedre kant
                var right = left + 55; // Höger kant (vänster + 55 mm)

                textFrame.geometricBounds = [top, left, bottom, right]; // Tilldela nya gränser
            }
        }
    }

    // Återställ GREP-inställningar
    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;
}


// Hantering för sex tal
function handleSixNumbers(textFrame) {
    // Nollställ eventuella gamla inställningar för GREP
    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;

    // Spara de nuvarande måttenhetsinställningarna
    var originalUnits = app.activeDocument.viewPreferences.horizontalMeasurementUnits;

    // Ställ in måttenheten till millimeter
    app.activeDocument.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.MILLIMETERS;
    app.activeDocument.viewPreferences.verticalMeasurementUnits = MeasurementUnits.MILLIMETERS;

    // Utför GREP-sökning och ersättning
    app.findGrepPreferences.findWhat = '\\s*(\\d+)\\s+(\\d+)\\s+(\\d+)\\s+(\\d+)\\s+(\\d+)\\s+(\\d+)\\s*';
    app.changeGrepPreferences.changeTo = '$1\\t\\t$2\\t$3\\t$4\\t$5 $6';
    textFrame.texts[0].changeGrep();

    // Sätt textrutans bredd till exakt 69.5 mm
    var bounds = textFrame.geometricBounds;
    bounds[3] = bounds[1] + 69.5; // Sätt högerkanten till vänsterkanten + 69.5 mm
    textFrame.geometricBounds = bounds;

    // Återställ måttenhetsinställningarna
    app.activeDocument.viewPreferences.horizontalMeasurementUnits = originalUnits;
    app.activeDocument.viewPreferences.verticalMeasurementUnits = originalUnits;

    // Återställ GREP-inställningarna
    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;
}

app.scriptPreferences.enableRedraw = true;
},
ScriptLanguage.JAVASCRIPT,
[],
UndoModes.FAST_ENTIRE_SCRIPT,
"AutomatisktHanteraTextrutor - Performance Boost"
);