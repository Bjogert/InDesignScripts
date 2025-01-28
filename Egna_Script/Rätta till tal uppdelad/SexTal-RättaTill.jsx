// TabbaSexTal-RensaOchFormatera-MedMillimeterFix.jsx
// -------------

// Nollställ eventuella gamla inställningar för GREP
app.findGrepPreferences = NothingEnum.nothing;
app.changeGrepPreferences = NothingEnum.nothing;

// Spara de nuvarande måttenhetsinställningarna
var originalUnits = app.activeDocument.viewPreferences.horizontalMeasurementUnits;

// Ställ in måttenheten till millimeter
app.activeDocument.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.MILLIMETERS;
app.activeDocument.viewPreferences.verticalMeasurementUnits = MeasurementUnits.MILLIMETERS;

// Sök efter alla sex tal separerade av valfria tabbar eller mellanslag
app.findGrepPreferences.findWhat = '\\s*(\\d+)\\s+(\\d+)\\s+(\\d+)\\s+(\\d+)\\s+(\\d+)\\s+(\\d+)\\s*';

// Byt ut enligt specifik formatering:
// - Ingen tabb före första talet.
// - Två tabbar före andra talet.
// - En tabb mellan tredje, fjärde och femte talet.
// - Ingen tabb framför sista talet.
app.changeGrepPreferences.changeTo = '$1\\t\\t$2\\t$3\\t$4\\t$5 $6';

// Tillämpa ändringen på alla markerade textramar
if (app.selection.length > 0) {
    for (var i = 0; i < app.selection.length; i++) {
        var textFrame = app.selection[i];

        // Kontrollera att objektet har textinnehåll
        if (textFrame.hasOwnProperty("texts")) {
            // Utför GREP-bytet
            textFrame.texts[0].changeGrep();

            // Sätt textrutans bredd till exakt 69 mm
            var bounds = textFrame.geometricBounds; // Hämta nuvarande geometriska gränser
            bounds[3] = bounds[1] + 69; // Sätt högerkanten till vänsterkanten + 69 mm
            textFrame.geometricBounds = bounds; // Tilldela nya gränser
        }
    }
} else {
    alert("Markera en eller flera textramar innan du kör skriptet.");
}

// Återställ måttenhetsinställningar
app.activeDocument.viewPreferences.horizontalMeasurementUnits = originalUnits;
app.activeDocument.viewPreferences.verticalMeasurementUnits = originalUnits;

// Återställ GREP-inställningar
app.findGrepPreferences = NothingEnum.nothing;
app.changeGrepPreferences = NothingEnum.nothing;
