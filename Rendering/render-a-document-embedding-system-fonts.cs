using System;
using Aspose.Words;
using Aspose.Words.Fonts;

class EmbedSystemFontsExample
{
    static void Main()
    {
        // Path to the folder where the output document will be saved.
        string outputPath = "EmbeddedSystemFonts.docx";

        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to add some text with a system font.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "MS Mincho"; // Example of a system font (Japanese).
        builder.Writeln("This document embeds the system font used above.");

        // Access the FontInfoCollection of the document.
        FontInfoCollection fontInfos = doc.FontInfos;

        // Enable embedding of TrueType fonts and system fonts.
        fontInfos.EmbedTrueTypeFonts = true;
        fontInfos.EmbedSystemFonts = true;

        // Optionally, decide whether to embed a subset of the fonts.
        // Here we embed the full font to ensure the system font is fully available.
        fontInfos.SaveSubsetFonts = false;

        // Save the document; the system font will be embedded.
        doc.Save(outputPath);
    }
}
