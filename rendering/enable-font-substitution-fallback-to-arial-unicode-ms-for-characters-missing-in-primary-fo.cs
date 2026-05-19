using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Create output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Build a document with a missing font and Unicode characters.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Font.Name = "NonExistentFont";
        builder.Writeln("This line uses a missing font and will be substituted.");

        // Unicode characters that may be missing in the primary font.
        builder.Writeln("Unicode test: 漢字, 😊, 𝔘𝔫𝔦𝔠𝔬𝔡𝔢");

        // Configure font settings to fall back to Arial Unicode MS.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial Unicode MS";
        doc.FontSettings = fontSettings;

        // Save the document as PDF.
        string pdfPath = Path.Combine(outputDir, "FallbackExample.pdf");
        doc.Save(pdfPath);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the PDF file.");

        Console.WriteLine($"PDF saved to: {pdfPath}");
    }
}
