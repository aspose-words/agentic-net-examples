using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths used in the example.
        string fontsFolder = "/usr/share/fonts/truetype"; // Linux system fonts folder.
        string outputPdf = "RenderedOutput.pdf";

        // Create a simple document with text that uses a font likely not installed by default.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "DejaVu Sans";
        builder.Writeln("This document is rendered using DejaVu Sans font on a headless Linux server.");

        // Configure FontSettings to look for fonts in the specified Linux fonts folder.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolder(fontsFolder, recursive: true);
        doc.FontSettings = fontSettings;

        // Render the document to PDF.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        doc.Save(outputPdf, pdfOptions);

        // Verify that the PDF was created successfully.
        if (!File.Exists(outputPdf))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPdf}");
        }

        // Optionally, inform that the process completed.
        Console.WriteLine($"Document rendered and saved to '{outputPdf}'.");
    }
}
