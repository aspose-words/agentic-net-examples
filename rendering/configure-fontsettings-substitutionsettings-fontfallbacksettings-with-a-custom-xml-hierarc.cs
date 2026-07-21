using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple XML file that defines custom font fallback rules.
        string fallbackXmlPath = Path.Combine(artifactsDir, "CustomFallback.xml");
        string fallbackXmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<FallbackSettings>
    <Range UnicodeRange=""0000-00FF"" Font=""Arial"" />
    <Range UnicodeRange=""0100-024F"" Font=""Times New Roman"" />
</FallbackSettings>";
        File.WriteAllText(fallbackXmlPath, fallbackXmlContent);

        // Create a new document with text that uses a font not present in the system.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "MissingFont";
        builder.Writeln("This text uses a missing font and will trigger fallback.");

        // Configure FontSettings and load the custom fallback definition.
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;
        fontSettings.FallbackSettings.Load(fallbackXmlPath);

        // Save the document to PDF – the fallback settings will be applied during rendering.
        string pdfPath = Path.Combine(artifactsDir, "Result.pdf");
        doc.Save(pdfPath);

        // Simple validation that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF output was not created.");

        // Optionally, save the effective fallback settings back to another file.
        string savedFallbackPath = Path.Combine(artifactsDir, "SavedFallback.xml");
        doc.FontSettings.FallbackSettings.Save(savedFallbackPath);
    }
}
