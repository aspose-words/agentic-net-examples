using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output directory and ensure it exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Path for the resulting PDF file.
        string pdfPath = Path.Combine(outputDir, "DocumentWithFallback.pdf");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a font that does not exist on the system to trigger fallback.
        builder.Font.Name = "NonExistentFont";
        builder.Writeln("This text uses a missing font and will be rendered with the fallback font.");

        // Configure font settings to use a specific fallback font.
        FontSettings fontSettings = new FontSettings();
        // The default font that will replace any missing TrueType font.
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Times New Roman";
        doc.FontSettings = fontSettings;

        // Save the document as PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException($"Failed to create PDF file at '{pdfPath}'.");
        }

        // Optionally, output the path of the generated file.
        Console.WriteLine($"PDF saved successfully to: {pdfPath}");
    }
}
