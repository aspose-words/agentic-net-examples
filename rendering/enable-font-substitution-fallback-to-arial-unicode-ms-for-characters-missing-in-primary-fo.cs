using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a font that is unlikely to be present on the system.
        builder.Font.Name = "MissingFont";

        // Write some text that contains characters outside the basic Latin range.
        // These characters will trigger font fallback.
        builder.Writeln("Hello world! Привет мир! こんにちは世界!");

        // Configure font settings to fall back to Arial Unicode MS when a font is missing.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial Unicode MS";
        doc.FontSettings = fontSettings;

        // Define the output PDF path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FallbackExample.pdf");

        // Save the document as PDF.
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        doc.Save(outputPath, saveOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // The program ends here without waiting for user input.
    }
}
