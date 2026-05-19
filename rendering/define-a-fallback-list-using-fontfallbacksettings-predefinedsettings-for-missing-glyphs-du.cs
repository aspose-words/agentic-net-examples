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

        // Create a new blank document.
        Document doc = new Document();

        // Set up custom font settings for the document.
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;

        // Load a predefined fallback scheme (Microsoft Office fallback).
        FontFallbackSettings fallback = fontSettings.FallbackSettings;
        fallback.LoadMsOfficeFallbackSettings();

        // Save the fallback scheme to an XML file (optional, demonstrates the Save method).
        string fallbackPath = Path.Combine(artifactsDir, "FallbackSettings.xml");
        fallback.Save(fallbackPath);

        // Build document content using a font that does not exist on the system.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Missing Font";
        builder.Writeln("This line uses a missing font. The fallback settings will supply a substitute.");

        // Render the document to PDF. The missing glyphs will be rendered using the fallback fonts.
        string pdfPath = Path.Combine(artifactsDir, "Output.pdf");
        doc.Save(pdfPath);

        // Simple validation that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the PDF output.");

        // Optionally, indicate success (no interactive prompts required).
        Console.WriteLine("Document rendered successfully. Files saved to: " + artifactsDir);
    }
}
