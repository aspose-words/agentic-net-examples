using System;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Build the document content with different fonts.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Writeln("This line uses the Arial font.");

        // Use a custom font that may not be installed on the system.
        builder.Font.Name = "Kreon-Regular";
        builder.Writeln("This line uses the Kreon-Regular font.");

        // Configure font settings so Aspose.Words can locate the custom font files.
        FontSettings fontSettings = new FontSettings();
        // Path to the folder that contains the required TrueType fonts.
        // Replace with the actual folder on your machine.
        string fontsFolder = @"C:\MyFonts";
        // Set the folder as a font source; true = search subfolders recursively.
        fontSettings.SetFontsFolder(fontsFolder, true);
        // Apply the font settings to the document.
        doc.FontSettings = fontSettings;

        // Render the first page to a PNG image using ImageSaveOptions (no System.Drawing required).
        ImageSaveOptions imgOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Render only the first page (zero‑based index).
            PageSet = new PageSet(0),
            // 100 DPI gives roughly 6x6 inches for a 600x600 pixel image.
            Resolution = 100
        };
        doc.Save(@"C:\Output\RenderedPage.png", imgOptions);

        // Save the document to PDF, embedding the fonts used in the document.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Embed the full font data so the PDF looks the same on any machine.
            EmbedFullFonts = true
        };
        doc.Save(@"C:\Output\Document.pdf", pdfOptions);
    }
}
