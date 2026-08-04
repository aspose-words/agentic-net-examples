using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define folders for output files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();

        // Set up font settings for the document.
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;

        // Load a predefined fallback scheme that mimics Microsoft Office.
        FontFallbackSettings fallback = fontSettings.FallbackSettings;
        fallback.LoadMsOfficeFallbackSettings();

        // (Optional) Save the loaded fallback settings to an XML file for inspection.
        string fallbackXmlPath = Path.Combine(artifactsDir, "FallbackSettings.xml");
        fallback.Save(fallbackXmlPath);

        // Write some text using a font that does not exist on the system.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Missing Font";
        builder.Writeln("This line uses a missing font and will be rendered with a fallback font.");

        // Render the document to PDF.
        string pdfPath = Path.Combine(artifactsDir, "RenderedWithFallback.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("The PDF output was not created.", pdfPath);
    }
}
