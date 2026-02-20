using System;
using Aspose.Words;
using Aspose.Words.Fonts;

class EmbedFontsExample
{
    static void Main()
    {
        // Path to the folder that contains the source document.
        string dataDir = @"C:\Data\";

        // Load an existing DOCX document.
        Document doc = new Document(dataDir + "input.docx");

        // Get the collection that controls font embedding for the document.
        FontInfoCollection fontInfos = doc.FontInfos;

        // Enable embedding of all TrueType fonts.
        fontInfos.EmbedTrueTypeFonts = true;

        // Enable embedding of system fonts (e.g., Arial, Times New Roman).
        fontInfos.EmbedSystemFonts = true;

        // Save a subset of each font rather than the whole font file (optional).
        fontInfos.SaveSubsetFonts = true;

        // Save the document with the embedded fonts.
        doc.Save(dataDir + "output.docx");
    }
}
