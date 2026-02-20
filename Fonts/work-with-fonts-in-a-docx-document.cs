using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Paths to the input document, output document and a folder that contains custom fonts.
        string dataDir = @"C:\Docs\";
        string inputFile = Path.Combine(dataDir, "Input.docx");
        string outputFile = Path.Combine(dataDir, "Output.docx");
        string fontsFolder = Path.Combine(dataDir, "MyFonts");

        // Load the existing DOCX document.
        Document doc = new Document(inputFile);

        // Configure font settings so Aspose.Words searches for fonts in the custom folder.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolder(fontsFolder, false);
        doc.FontSettings = fontSettings;

        // Enable embedding of all fonts (TrueType and system) used in the document.
        FontInfoCollection fontInfos = doc.FontInfos;
        fontInfos.EmbedTrueTypeFonts = true;
        fontInfos.EmbedSystemFonts = true;
        fontInfos.SaveSubsetFonts = true; // Save only the characters that are actually used.

        // List all fonts that the document references.
        Console.WriteLine("Fonts used in the document:");
        for (int i = 0; i < fontInfos.Count; i++)
        {
            FontInfo info = fontInfos[i];
            Console.WriteLine($"- Name: {info.Name}");
            Console.WriteLine($"  TrueType: {info.IsTrueType}");
            Console.WriteLine($"  Family: {info.Family}");
            Console.WriteLine($"  Pitch: {info.Pitch}");
        }

        // Save the document with the embedded fonts.
        doc.Save(outputFile, SaveFormat.Docx);
    }
}
