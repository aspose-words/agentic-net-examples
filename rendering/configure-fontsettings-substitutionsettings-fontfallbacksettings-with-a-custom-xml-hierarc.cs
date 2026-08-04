using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Prepare output directories
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for temporary XML files and final PDF
        string defaultFallbackPath = Path.Combine(outputDir, "default_fallback.xml");
        string customFallbackPath = Path.Combine(outputDir, "custom_fallback.xml");
        string pdfPath = Path.Combine(outputDir, "Result.pdf");

        // Create a new empty document
        Document doc = new Document();

        // Create FontSettings and associate it with the document
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;

        // Build an automatic fallback scheme and save it to an XML file
        FontFallbackSettings fallback = fontSettings.FallbackSettings;
        fallback.BuildAutomatic();
        fallback.Save(defaultFallbackPath);

        // Load the generated XML, modify it to create a custom hierarchy, and save as a new file
        string xmlContent = File.ReadAllText(defaultFallbackPath);
        // Example modification: replace any occurrence of a font name with "Arial"
        // (In a real scenario you would adjust the XML structure as needed)
        xmlContent = xmlContent.Replace("AllegroOpen", "Arial");
        File.WriteAllText(customFallbackPath, xmlContent);

        // Load the custom fallback settings from the modified XML
        fallback.Load(customFallbackPath);

        // Add some text using a font that likely does not exist to trigger fallback
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Missing Font";
        builder.Writeln("This text uses a missing font and will be rendered using the fallback settings.");

        // Save the document as PDF to verify that the fallback settings are applied
        doc.Save(pdfPath);

        // Simple validation that the PDF was created
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF output was not created.");

        // Cleanup temporary XML files (optional)
        File.Delete(defaultFallbackPath);
        File.Delete(customFallbackPath);
    }
}
