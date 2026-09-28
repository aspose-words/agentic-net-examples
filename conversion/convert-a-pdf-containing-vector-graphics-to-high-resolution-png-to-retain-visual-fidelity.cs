using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample PDF that contains vector graphics.
        const string pdfPath = "sample.pdf";
        CreateSamplePdf(pdfPath);

        // Step 2: Load the PDF document.
        LoadOptions loadOptions = new LoadOptions { LoadFormat = LoadFormat.Pdf };
        Document pdfDocument = new Document(pdfPath, loadOptions);

        // Step 3: Convert the first page of the PDF to a high‑resolution PNG.
        const string pngPath = "output.png";
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            Resolution = 300,                     // High DPI for visual fidelity.
            PageSet = new PageSet(0)              // Convert only the first page.
        };
        pdfDocument.Save(pngPath, pngOptions);

        // Step 4: Validate that the PNG file was created and contains data.
        if (!File.Exists(pngPath) || new FileInfo(pngPath).Length == 0)
            throw new InvalidOperationException("PNG conversion failed: output file is missing or empty.");
    }

    private static void CreateSamplePdf(string outputPath)
    {
        // Create a simple Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text.
        builder.Writeln("Sample PDF with vector graphics.");

        // Insert a vector shape (a star) to ensure vector content.
        Shape star = builder.InsertShape(ShapeType.Star, 200, 200);
        // No explicit color is set to avoid System.Drawing usage.

        // Save the document as PDF.
        doc.Save(outputPath, SaveFormat.Pdf);
    }
}
