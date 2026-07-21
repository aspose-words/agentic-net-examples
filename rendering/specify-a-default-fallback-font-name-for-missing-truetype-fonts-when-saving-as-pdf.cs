using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a font that does not exist on the system.
        builder.Font.Name = "NonExistentFont";
        builder.Writeln("This text uses a missing font and will be substituted.");

        // Configure font substitution: set a default fallback font.
        FontSettings fontSettings = new FontSettings();
        DefaultFontSubstitutionRule defaultSubstitution = fontSettings.SubstitutionSettings.DefaultFontSubstitution;
        defaultSubstitution.DefaultFontName = "Courier New"; // fallback font
        doc.FontSettings = fontSettings;

        // Save the document as PDF.
        string pdfPath = Path.Combine(outputDir, "DocumentWithFallback.pdf");
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // Optionally, you could inspect the PDF bytes for embedded font markers here.
        // For this example we simply finish execution.
    }
}
