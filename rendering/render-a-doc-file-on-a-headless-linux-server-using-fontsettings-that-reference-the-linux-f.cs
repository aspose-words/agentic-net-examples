using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Prepare output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string pdfPath = Path.Combine(outputDir, "RenderedDocument.pdf");

        // Create a simple document with a font that exists on Linux
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "DejaVu Sans";
        builder.Writeln("This is a test document rendered on Linux using DejaVu Sans font.");

        // Configure FontSettings to point to the Linux fonts folder
        FontSettings fontSettings = new FontSettings();
        string linuxFontsFolder = "/usr/share/fonts/truetype";
        fontSettings.SetFontsFolder(linuxFontsFolder, true);
        doc.FontSettings = fontSettings;

        // Render the document to PDF
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException($"Failed to create PDF at '{pdfPath}'.");
        }

        // Indicate success
        Console.WriteLine($"PDF rendered successfully: {pdfPath}");
    }
}
