using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a temporary folder for custom fonts.
        string fontsFolder = Path.Combine(Path.GetTempPath(), "CustomFonts");
        Directory.CreateDirectory(fontsFolder);

        // Create a new document and add a paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This text will use fonts from the custom fonts folder.");

        // Create FontSettings and point it to the custom fonts folder.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolder(fontsFolder, false);

        // Assign the FontSettings instance to the document.
        doc.FontSettings = fontSettings;

        // Define output file path.
        string outputPath = Path.Combine(Path.GetTempPath(), "RenderedDocument.pdf");

        // Render the document to PDF.
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        doc.Save(outputPath, saveOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Clean up temporary resources (optional).
        // Directory.Delete(fontsFolder, true);
    }
}
