using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output folder and file.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(outputDir);
        string pdfPath = Path.Combine(outputDir, "DocumentWithFallback.pdf");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a font that does not exist on the system to trigger fallback.
        builder.Font.Name = "NonExistentFont";
        builder.Writeln("This text uses a missing font and will be rendered with the fallback font.");

        // Configure font settings to use a specific fallback font.
        FontSettings fontSettings = new FontSettings();
        // The DefaultFontSubstitution rule defines the font used when the requested one is unavailable.
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Courier New";
        doc.FontSettings = fontSettings;

        // Save the document as PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("The PDF file was not created.", pdfPath);

        // Optional: output the path for confirmation (no interactive input required).
        Console.WriteLine($"PDF saved to: {pdfPath}");
    }
}
