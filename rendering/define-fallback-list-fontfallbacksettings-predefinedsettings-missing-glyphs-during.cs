using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

class FontFallbackExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize FontSettings for the document.
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;

        // Access the fallback settings.
        FontFallbackSettings fallbackSettings = fontSettings.FallbackSettings;

        // Load a predefined fallback scheme that mimics Microsoft Office.
        // This uses the built‑in LoadMsOfficeFallbackSettings rule.
        fallbackSettings.LoadMsOfficeFallbackSettings();

        // (Optional) Save the loaded fallback scheme to an XML file for inspection.
        // Uses the Save(string) rule.
        string fallbackPath = Path.Combine(Environment.CurrentDirectory, "MsOfficeFallback.xml");
        fallbackSettings.Save(fallbackPath);

        // Use a font that does not exist in any of the font sources.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Missing Font";
        builder.Writeln("This text uses a missing font and will be rendered using the fallback scheme.");

        // Save the document to PDF (or any other supported format).
        string outputPath = Path.Combine(Environment.CurrentDirectory, "FallbackDemo.pdf");
        doc.Save(outputPath);
    }
}
