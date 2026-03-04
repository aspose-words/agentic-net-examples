using System;
using Aspose.Words;
using Aspose.Words.Fonts;

class LoadSystemFontsExample
{
    static void Main()
    {
        // Path to the source DOCX document.
        string inputPath = @"C:\Docs\Sample.docx";

        // Path where the processed document will be saved.
        string outputPath = @"C:\Docs\Sample_WithSystemFonts.pdf";

        // Load the document.
        Document doc = new Document(inputPath);

        // Create a new FontSettings instance.
        FontSettings fontSettings = new FontSettings();

        // Retrieve all system font folders.
        string[] systemFontFolders = SystemFontSource.GetSystemFontFolders();

        // Configure the FontSettings to use the system font folders.
        // The second argument (true) enables recursive scanning of subfolders.
        fontSettings.SetFontsFolders(systemFontFolders, true);

        // Assign the configured FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Save the document (PDF format in this example).
        doc.Save(outputPath);
    }
}
