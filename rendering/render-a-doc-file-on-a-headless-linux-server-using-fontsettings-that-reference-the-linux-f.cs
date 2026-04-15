using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Prepare output directory and file path.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "Rendered.pdf");

        // Create a simple document with some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "DejaVu Sans";
        builder.Font.Size = 24;
        builder.Writeln("Hello from Aspose.Words on Linux!");
        builder.Writeln("Rendering with custom FontSettings.");

        // Configure FontSettings to use the Linux fonts folder.
        FontSettings fontSettings = new FontSettings();
        // Typical location of TrueType fonts on Linux.
        string linuxFontsFolder = "/usr/share/fonts/truetype";
        fontSettings.SetFontsFolder(linuxFontsFolder, recursive: true);
        doc.FontSettings = fontSettings;

        // Render the document to PDF.
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create output file: {outputPath}");
        }

        // Confirmation (optional, non‑interactive).
        Console.WriteLine($"Document rendered successfully to: {outputPath}");
    }
}
