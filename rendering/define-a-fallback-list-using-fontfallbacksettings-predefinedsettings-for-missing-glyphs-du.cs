using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

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
        builder.Font.Name = "MissingFont";
        builder.Writeln("This text uses a missing font. The glyphs should be rendered using fallback fonts.");

        // Configure font settings with predefined Microsoft Office fallback scheme.
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;
        FontFallbackSettings fallback = fontSettings.FallbackSettings;
        fallback.LoadMsOfficeFallbackSettings();

        // Save the fallback configuration for reference.
        string fallbackPath = Path.Combine(outputDir, "FallbackSettings.xml");
        fallback.Save(fallbackPath);

        // Render the document to PDF.
        string pdfPath = Path.Combine(outputDir, "DocumentWithFallback.pdf");
        doc.Save(pdfPath);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the PDF output.");

        Console.WriteLine($"PDF saved to: {pdfPath}");
        Console.WriteLine($"Fallback settings saved to: {fallbackPath}");
    }
}
