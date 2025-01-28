(function () {
    // 1) Kontrollera om dokument är öppet
    if (app.documents.length === 0) {
        alert("No document is open.");
        return;
    }

    var myDocument = app.activeDocument;

    // 2) Definiera varianter för export
    var exportVariants = [
        {
            pdfName: "Decibel-SEK",
            layersToShow: ["Lager 1", "Svenska", "SEK_SE"]
        },
        {
            pdfName: "Decibel-NOK",
            layersToShow: ["Lager 1", "Svenska", "NOK_NO"]
        },
        {
            pdfName: "Decibel-DKK",
            layersToShow: ["Lager 1", "Svenska", "DKK_DK"]
        },
        {
            pdfName: "Decibel-EURO",
            layersToShow: ["Lager 1", "Engelska", "EUR_EU"]
        },
        {
            pdfName: "Decibel-GBP",
            layersToShow: ["Lager 1", "Engelska", "GBP_UK"]
        },
        {
            pdfName: "Decibel-EURO Europe",
            layersToShow: ["Lager 1", "Engelska", "EUR_BEIR", "Villkor - Euro Europe BeLuIr engelska"]
        },
        {
            pdfName: "Decibel-EURO Switzerland (Eng)",
            layersToShow: ["Lager 1", "Engelska", "EUR_BEIR", "Villkor - Schweiz engelska"]
        },
        {
            pdfName: "Decibel-EURO DE Deutschland",
            layersToShow: ["Lager 1", "Tyska", "EUR_DE", "Villkor - Tyska Marcus"]
        },
        {
            pdfName: "Decibel-EURO Österreich Switzerland (DE)",
            layersToShow: ["Lager 1", "Tyska", "EUR_DE", "Villkor - Schweiz Österrike tyska"]
        },
        {
            pdfName: "Decibel-EURO The Netherlands",
            layersToShow: ["Lager 1", "Engelska", "EUR_EURAEA", "Villkor - Nederländerna"]
        },
        {
            pdfName: "Decibel-EURO France Agence Pise",
            layersToShow: ["Lager 1", "Engelska", "EUR_EURAEA", "Villkor - Franska"]
        },
        {
            pdfName: "Decibel-USD",
            layersToShow: ["Lager 1", "Engelska INCHES", "USD_US"]
        }
    ];

    // 3) Ange var PDF-filerna ska sparas
    var outputFolder = Folder("C:/Users/robert/OneDrive - Johanson Design AB/Skrivbordet/Prislistor_Decibel 2025 Lokala");
    if (!outputFolder.exists) {
        alert("Output folder does not exist.");
        return;
    }

    // 4) Skapa en loggfil
    var logFile = File(outputFolder.fsName + "/ErrorLog.txt");
    logFile.open("w");
    logFile.writeln("LOGG - " + new Date());
    logFile.writeln("--------------------------");

    // 5) Sätt inställningar för Interaktiv PDF
    with (app.interactivePDFExportPreferences) {
        pageRange = PageRange.ALL_PAGES;
        exportReaderSpreads = false;
        rasterizePages = false;
        generateThumbnails = true;
        pdfRasterCompression = PDFRasterCompressionOptions.JPEG_COMPRESSION;
        pdfJPEGQuality = PDFJPEGQualityOptions.HIGH;
        pdfImageResolution = 300;
    }

    // 6) Funktion för att hitta överskjuten *text* (ignorerar tomma/bildramar)
    function oversetStoriesWithText(doc) {
        var overs = [];
        for (var i = 0; i < doc.stories.length; i++) {
            var story = doc.stories[i];
            // Kolla om storyn är overset och faktiskt har text (characters.length > 0)
            if (story.overflows && story.characters.length > 0) {
                overs.push({
                    index: i,
                    name: story.name || "Story utan namn"
                });
            }
        }
        return overs;
    }

    var allLayers = myDocument.layers;

    // 7) Loopa igenom alla varianter
    for (var i = 0; i < exportVariants.length; i++) {
        var variant = exportVariants[i];

        // Släck alla lager
        for (var l = 0; l < allLayers.length; l++) {
            allLayers[l].visible = false;
        }

        // Tänd enbart de lager som behövs
        for (var k = 0; k < variant.layersToShow.length; k++) {
            try {
                myDocument.layers.itemByName(variant.layersToShow[k]).visible = true;
            } catch (err) {
                logFile.writeln(
                    "VARNING: Saknat lager '" + variant.layersToShow[k] + 
                    "' i variant '" + variant.pdfName + "'"
                );
            }
        }

        // Kolla overset (text)
        var overset = oversetStoriesWithText(myDocument);

        // Bygg filnamn
        var finalName = variant.pdfName + ".pdf";
        var pdfFile = File(outputFolder + "/" + finalName);

        // Exportera Interaktiv PDF
        myDocument.exportFile(ExportFormat.INTERACTIVE_PDF, pdfFile, false);

        // Logga resultat
        if (overset.length > 0) {
            logFile.writeln("OVERSIZED TEXT i '" + variant.pdfName + "':");
            for (var x = 0; x < overset.length; x++) {
                logFile.writeln(
                    " - StoryIndex=" + overset[x].index + 
                    ", Namn=\"" + overset[x].name + "\""
                );
            }
        } else {
            logFile.writeln("OK: " + variant.pdfName);
        }
    }

    // 8) Avsluta loggen
    logFile.writeln("--------------------------");
    logFile.writeln("KLART");
    logFile.close();

    alert("WOW! Fäääädig! Se ErrorLog.txt för information.");
})();
