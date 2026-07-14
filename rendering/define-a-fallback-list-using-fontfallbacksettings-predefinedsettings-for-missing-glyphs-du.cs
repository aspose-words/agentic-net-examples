using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define an output folder and ensure it exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Path for the rendered PDF file.
        string pdfPath = Path.Combine(outputDir, "FallbackDemo.pdf");

        // Create a new blank document.
        Document doc = new Document();

        // Create FontSettings and assign it to the document.
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;

        // Load a predefined fallback scheme (Microsoft Office fallback).
        // This scheme defines which fonts to use when the original font lacks required glyphs.
        fontSettings.FallbackSettings.LoadMsOfficeFallbackSettings();

        // Build the document content using a font that does not exist on the system.
        // The missing glyphs will be rendered using the fallback fonts defined above.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "MissingFont";
        builder.Writeln("This line uses a missing font. Characters that are not present will fall back:");
        // Example Unicode characters that typically require fallback (Telugu block).
        builder.Writeln("\u0C05\u0C06\u0C07");

        // Render the document to PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the PDF file.");

        // Inform the user where the file was saved.
        Console.WriteLine($"PDF successfully saved to: {pdfPath}");
    }
}
