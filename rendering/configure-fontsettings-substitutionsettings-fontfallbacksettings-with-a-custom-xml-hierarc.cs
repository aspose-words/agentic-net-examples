using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Prepare output folders.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Path for the custom fallback XML and the resulting PDF.
        string fallbackXmlPath = Path.Combine(artifactsDir, "CustomFontFallback.xml");
        string pdfPath = Path.Combine(artifactsDir, "Result.pdf");

        // Create a new empty document.
        Document doc = new Document();

        // Create FontSettings and obtain its fallback settings object.
        FontSettings fontSettings = new FontSettings();
        FontFallbackSettings fallbackSettings = fontSettings.FallbackSettings;

        // Build an automatic fallback scheme and save it as a custom XML file.
        fallbackSettings.BuildAutomatic();
        fallbackSettings.Save(fallbackXmlPath);

        // Load the custom fallback settings from the XML file.
        fallbackSettings.Load(fallbackXmlPath);

        // Assign the configured FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Write some text using a font that is unlikely to be present to trigger fallback.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "MissingFont";
        builder.Writeln("This text uses a missing font and will be rendered using the fallback scheme.");

        // Save the document as PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Optionally, indicate success (no console input required).
        Console.WriteLine("PDF generated successfully at: " + pdfPath);
    }
}
