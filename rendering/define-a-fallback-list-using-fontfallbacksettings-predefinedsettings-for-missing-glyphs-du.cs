using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create an output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document.
        Document doc = new Document();

        // Set up font settings for the document.
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;

        // Load the predefined Microsoft Office fallback scheme.
        FontFallbackSettings fallback = fontSettings.FallbackSettings;
        fallback.LoadMsOfficeFallbackSettings();

        // Save the fallback configuration to an XML file (optional, for inspection).
        string fallbackPath = Path.Combine(outputDir, "FallbackSettings.xml");
        fallback.Save(fallbackPath);

        // Build document content using a font that does not exist.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Missing Font";
        builder.Writeln("This text uses a missing font. The following characters require fallback:");
        builder.Writeln("Latin: ABC");
        builder.Writeln("Cyrillic: АБВ");
        builder.Writeln("Greek: ΑΒΓ");
        builder.Writeln("Arabic: أبج");
        builder.Writeln("Hebrew: אבג");
        builder.Writeln("Chinese: 汉字");
        builder.Writeln("Emoji: 😊");

        // Render the document to PDF, which will apply the fallback settings.
        string pdfPath = Path.Combine(outputDir, "DocumentWithFallback.pdf");
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        doc.Save(pdfPath, pdfOptions);

        // Simple validation that the files were created.
        if (!File.Exists(fallbackPath))
            throw new Exception("Fallback settings file was not created.");
        if (!File.Exists(pdfPath))
            throw new Exception("PDF output file was not created.");
    }
}
