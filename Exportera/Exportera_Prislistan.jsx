(function () {
    if (app.documents.length === 0) {
        alert("Inget dokument är öppet.");
        return;
    }

    var myDocument = app.activeDocument;

    var exportVariants = [
        {
            pdfName: "SEK",
            layersToShow: ["Lager 1", "Svenska", "SEK_SE"] 
        },
        {
            pdfName: "NOK",
            layersToShow: ["Lager 1", "Svenska", "SEK_NO"]
        },
        {
            pdfName: "DKK",
            layersToShow: ["Lager 1", "Svenska", "DKK_DK"]
        },
        {
            pdfName: "EURO",
            layersToShow: ["Lager 1", "Engelska", "EUR_EU"]
        },
        {
            pdfName: "GBP",
            layersToShow: ["Lager 1", "Engelska", "GBP_UK"]
        },
        {
            pdfName: "EURO Europe",
            layersToShow: ["Lager 1", "Engelska", "EUR_BEIR", "Villkor - Euro Europe BeLuIr engelska"]
        },
        {
            pdfName: "EURO Switzerland (Eng)",
            layersToShow: ["Lager 1", "Engelska", "EUR_BEIR", "Villkor - Schweiz engelska"]
        },
        {
            pdfName: "EURO DE Deutschland",
            layersToShow: ["Lager 1", "Tyska", "EUR_DE", "Villkor - Tyska Marcus"]
        },
        {
            pdfName: "EURO Österreich Switzerland (DE)",
            layersToShow: ["Lager 1", "Tyska", "EUR_DE", "Villkor - Schweiz Österrike tyska"]
        },
        {
            pdfName: "EURO The Netherlands",
            layersToShow: ["Lager 1", "Engelska", "EUR_EURAEA", "Villkor - Nederländerna"]
        },
        {
            pdfName: "EURO France Agence Pise",
            layersToShow: ["Lager 1", "Engelska", "EUR_EURAEA", "Villkor - Franska"]
        },
        {
            pdfName: "USD",
            layersToShow: ["Lager 1", "Engelska INCHES", "USD_US"]
        }
    ];

    var outputFolder = Folder("C:/Users/robert/OneDrive - Johanson Design AB/Skrivbordet/Prislistor 2025 Lokala");
    if (!outputFolder.exists) {
        alert("Sökvägen existerar inte.");
        return;
    }

    var logFile = File(outputFolder.fsName + "/ErrorLog.txt");
    logFile.open("w");
    logFile.writeln("LOGG - " + new Date());
    logFile.writeln("--------------------------");

    // Ställ in interaktiv PDF-export
    with (app.interactivePDFExportPreferences) {
        pageRange = PageRange.ALL_PAGES;
        exportReaderSpreads = false;
        rasterizePages = false;
        generateThumbnails = true;
        pdfRasterCompression = PDFRasterCompressionOptions.JPEG_COMPRESSION;
        pdfJPEGQuality = PDFJPEGQualityOptions.MEDIUM;
        pdfImageResolution = 144;
    }

    // Ny funktion som bara rapporterar storyn om den är overset *och* har text
    function oversetStoriesWithText(doc) {
        var overs = [];
        for (var i = 0; i < doc.stories.length; i++) {
            var st = doc.stories[i];
            if (st.overflows && st.characters.length > 0) {
                overs.push({
                    index: i,
                    name: st.name || "Story utan namn"
                });
            }
        }
        return overs;
    }

    var allLayers = myDocument.layers;

    // Exportloop
    for (var i = 0; i < exportVariants.length; i++) {
        var variant = exportVariants[i];

        // Släck alla lager
        for (var l = 0; l < allLayers.length; l++) {
            allLayers[l].visible = false;
        }

        // Försök tända "Vita boxar..." om det finns
        try {
            myDocument.layers.itemByName("Vita boxar som täcker siffror").visible = true;
        } catch (e) {
            logFile.writeln("VARNING: Lagret 'Vita boxar som täcker siffror' saknas.");
        }

        // Tänd de lager som krävs för denna variant
        for (var k = 0; k < variant.layersToShow.length; k++) {
            try {
                myDocument.layers.itemByName(variant.layersToShow[k]).visible = true;
            } catch (err) {
                logFile.writeln("VARNING: Saknat lager '" + variant.layersToShow[k] + 
                                "' i variant '" + variant.pdfName + "'");
            }
        }

        // Kolla overset (text)
        var overset = oversetStoriesWithText(myDocument);

        // Bygg filnamn (om overset finns: lägg till "_ErrorText")
        var finalName = variant.pdfName;
        if (overset.length > 0) {
            finalName += "_ErrorText";
        }
        finalName += ".pdf";

        var pdfFile = File(outputFolder + "/" + finalName);

        // Exportera Interaktiv PDF
        myDocument.exportFile(ExportFormat.INTERACTIVE_PDF, pdfFile, false);

        // Logga resultat
        if (overset.length > 0) {
            logFile.writeln("OVERSIZED TEXT i '" + variant.pdfName + "':");
            for (var x = 0; x < overset.length; x++) {
                logFile.writeln(" - StoryIndex=" + overset[x].index + 
                                ", Namn=\"" + overset[x].name + "\"");
            }
        } else {
            logFile.writeln("OK: " + variant.pdfName);
        }
    }

    logFile.writeln("--------------------------");
    logFile.writeln("KLART");
    logFile.close();

    alert("Färdig! Se ErrorLog.txt för information.");
})();
