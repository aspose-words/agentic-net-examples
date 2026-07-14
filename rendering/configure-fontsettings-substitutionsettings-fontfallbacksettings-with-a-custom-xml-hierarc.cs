using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Missing Font"; // Font that does not exist in the system.
        builder.Writeln("This text uses a missing font and will trigger fallback.");

        // Configure FontSettings with custom fallback settings.
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;
        FontFallbackSettings fallback = fontSettings.FallbackSettings;

        // Build an automatic fallback scheme and save it to a custom XML file.
        fallback.BuildAutomatic();
        string fallbackXmlPath = Path.Combine(artifactsDir, "CustomFallback.xml");
        fallback.Save(fallbackXmlPath);

        // Load the custom fallback settings from the XML file.
        fallback.Load(fallbackXmlPath);

        // Render the document to PDF.
        string pdfPath = Path.Combine(artifactsDir, "Output.pdf");
        doc.Save(pdfPath);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the PDF output.");

        // Optionally, save the currently used fallback settings for inspection.
        string savedFallbackPath = Path.Combine(artifactsDir, "SavedFallback.xml");
        fallback.Save(savedFallbackPath);
    }
}
