using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a folder for all generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a font that is unlikely to exist on the system to trigger fallback.
        builder.Font.Name = "MissingFont";
        builder.Writeln("This text uses a missing font. The fallback list will provide glyphs.");

        // Configure font settings for the document.
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;

        // Load a predefined fallback scheme (Microsoft Office fallback).
        FontFallbackSettings fallback = fontSettings.FallbackSettings;
        fallback.LoadMsOfficeFallbackSettings();

        // Save the fallback settings to an XML file (optional, for inspection).
        string fallbackPath = Path.Combine(artifactsDir, "FallbackSettings.xml");
        fallback.Save(fallbackPath);

        // Render the document to PDF using the fallback settings.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        string pdfPath = Path.Combine(artifactsDir, "DocumentWithFallback.pdf");
        doc.Save(pdfPath, pdfOptions);

        // Simple validation that the expected files were created.
        if (!File.Exists(fallbackPath))
            throw new FileNotFoundException("Fallback settings file was not created.", fallbackPath);
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF output file was not created.", pdfPath);
    }
}
