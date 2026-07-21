using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Path to the Linux TrueType fonts folder (common location). Adjust if necessary.
        const string linuxFontsFolder = "/usr/share/fonts/truetype";

        // Configure FontSettings to use the Linux fonts folder recursively.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolder(linuxFontsFolder, recursive: true);

        // Create a simple document and apply a font that is expected to be present in the Linux fonts folder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "DejaVu Sans"; // Example font available on many Linux distributions.
        builder.Writeln("Hello from Aspose.Words on a headless Linux server!");
        builder.Writeln("This document uses fonts from the Linux font directory.");

        // Assign the configured FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Render the document to PDF.
        const string outputFile = "RenderedDocument.pdf";
        doc.Save(outputFile, SaveFormat.Pdf);

        // Verify that the PDF was created successfully.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException($"Failed to create the output file: {outputFile}");

        // Optional: indicate success (no interactive prompts required).
        Console.WriteLine($"Document rendered successfully to '{outputFile}'.");
    }
}
