using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a font that is unlikely to be present on the system.
        builder.Font.Name = "Missing Font";
        builder.Writeln("This line uses a missing font and will be substituted.");

        // Add some Unicode characters that are not covered by most fonts.
        builder.Writeln("Unicode test: 漢字, 😊, 𝔘𝔫𝔦𝔠𝔬𝔡𝔢");

        // Configure font settings to fall back to Arial Unicode MS for any missing font.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial Unicode MS";
        doc.FontSettings = fontSettings;

        // Define output path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "FontFallbackExample.pdf");

        // Save the document to PDF.
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the PDF output file.");

        // Optionally, you could inspect the PDF for embedded font markers here.
        // For this example we simply complete execution.
    }
}
