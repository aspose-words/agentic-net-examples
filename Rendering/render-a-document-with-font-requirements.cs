using System;
using Aspose.Words;
using Aspose.Words.Fonts;

class RenderDocumentWithFonts
{
    static void Main()
    {
        // Paths to the data folder, custom fonts folder and output file.
        string dataDir = @"C:\Data\";
        string fontsDir = System.IO.Path.Combine(dataDir, "Fonts");
        string outputPath = System.IO.Path.Combine(dataDir, "RenderedDocument.pdf");

        // Create a new blank document.
        Document doc = new Document();

        // Build the document content with specific fonts.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Amethysta";               // Font that may not be installed on the system.
        builder.Writeln("The quick brown fox jumps over the lazy dog.");
        builder.Font.Name = "Arial";                  // Common system font.
        builder.Writeln("Hello world!");

        // Configure FontSettings to look for fonts in a custom folder.
        FontSettings fontSettings = new FontSettings();
        // 'true' enables recursive search in subfolders.
        fontSettings.SetFontsFolder(fontsDir, true);
        doc.FontSettings = fontSettings;

        // Save the document to PDF using the standard Save method.
        doc.Save(outputPath);
    }
}
