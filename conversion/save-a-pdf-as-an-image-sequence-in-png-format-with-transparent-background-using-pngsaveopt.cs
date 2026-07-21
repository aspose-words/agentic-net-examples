using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with three pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Page 1 content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3 content.");

        // Save the document as PDF (required step before conversion).
        const string pdfPath = "sample.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // Load the PDF for conversion.
        Document pdfDoc = new Document(pdfPath);

        // Prepare ImageSaveOptions for PNG with a transparent background.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Use System.Drawing.Color for the PaperColor property (required by the API).
            PaperColor = System.Drawing.Color.Transparent
        };

        // Export each page to a separate PNG file.
        for (int i = 0; i < pdfDoc.PageCount; i++)
        {
            pngOptions.PageSet = new PageSet(i); // Zero‑based page index.
            string pngPath = $"page_{i + 1}.png";
            pdfDoc.Save(pngPath, pngOptions);

            if (!File.Exists(pngPath))
                throw new InvalidOperationException($"PNG file '{pngPath}' was not created.");
        }

        Console.WriteLine("PDF successfully converted to PNG images with transparent background.");
    }
}
