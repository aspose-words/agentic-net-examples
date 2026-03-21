using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Loading;

class FontPriorityExample
{
    static void Main()
    {
        // Create a temporary folder that will act as the user‑specified font directory.
        string userFontsFolder = Path.Combine(Path.GetTempPath(), "MyFonts");
        Directory.CreateDirectory(userFontsFolder);

        // Create a folder font source for the user folder.
        // The third argument sets the priority – a higher value means higher priority.
        FolderFontSource userFontSource = new FolderFontSource(userFontsFolder, true, 1);

        // Obtain the default system font source(s). By default there is one SystemFontSource.
        FontSourceBase[] systemSources = FontSettings.DefaultInstance.GetFontsSources();

        // Combine the user font source with the system sources, placing the user source first.
        FontSourceBase[] combinedSources = new FontSourceBase[systemSources.Length + 1];
        combinedSources[0] = userFontSource;
        Array.Copy(systemSources, 0, combinedSources, 1, systemSources.Length);

        // Create a FontSettings instance and assign the combined sources.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsSources(combinedSources);

        // Apply the FontSettings while loading the document.
        LoadOptions loadOptions = new LoadOptions { FontSettings = fontSettings };

        // Create a simple document in memory (no external file needed).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, world! This document uses font priority settings.");

        // Save the document to PDF in the temporary folder.
        string outputPath = Path.Combine(Path.GetTempPath(), "Output.pdf");
        doc.Save(outputPath, SaveFormat.Pdf);

        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
