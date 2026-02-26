using System;
using Aspose.Words;
using Aspose.Words.Fonts;

class EmbedFontsExample
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Data\Document.docx";

        // Path where the DOCX with embedded fonts will be saved.
        string outputPath = @"C:\Output\EmbeddedFonts.docx";

        // Load the existing document.
        Document doc = new Document(inputPath);

        // Access the collection of fonts used in the document.
        FontInfoCollection fontInfos = doc.FontInfos;

        // Enable embedding of all TrueType fonts.
        fontInfos.EmbedTrueTypeFonts = true;

        // Also embed system fonts (required when EmbedTrueTypeFonts is true).
        fontInfos.EmbedSystemFonts = true;

        // Save a subset of each embedded font (only the glyphs used in the document).
        fontInfos.SaveSubsetFonts = true;

        // Save the document with the embedded fonts.
        doc.Save(outputPath);
    }
}
