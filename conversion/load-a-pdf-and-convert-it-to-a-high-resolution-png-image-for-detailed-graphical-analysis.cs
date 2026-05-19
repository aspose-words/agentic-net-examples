using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary PDF and the resulting PNG.
        const string pdfPath = "sample.pdf";
        const string pngPath = "sample_high_res.png";

        // -----------------------------------------------------------------
        // Create a sample Word document and save it as a PDF.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample PDF document for conversion.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page content.");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Expected PDF file was not created.");

        // -----------------------------------------------------------------
        // Load the PDF and convert it to a high‑resolution PNG image.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // Configure image save options for high quality.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            Resolution = 300,                 // 300 DPI for high resolution.
            UseHighQualityRendering = true   // Enable high‑quality rendering algorithms.
        };

        // Save the first page of the PDF as a PNG image.
        pdfDoc.Save(pngPath, pngOptions);

        // Verify that the PNG file was created.
        if (!File.Exists(pngPath))
            throw new InvalidOperationException("Expected PNG file was not created.");

        // Indicate successful conversion.
        Console.WriteLine("PDF successfully converted to high‑resolution PNG.");
    }
}
