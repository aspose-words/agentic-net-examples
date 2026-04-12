using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string pdfPath = Path.Combine(Directory.GetCurrentDirectory(), "sample_vector.pdf");
        string pngPath = Path.Combine(Directory.GetCurrentDirectory(), "sample_vector.png");

        // -----------------------------------------------------------------
        // Step 1: Create a sample PDF that contains vector graphics.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Sample PDF with vector graphics");
        // Insert a rectangle shape – this is stored as a vector graphic.
        builder.InsertShape(ShapeType.Rectangle, 200, 100);

        // Save the document as PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
            throw new InvalidOperationException("Failed to create the sample PDF.");

        // -----------------------------------------------------------------
        // Step 2: Load the PDF and convert the first page to a high‑resolution PNG.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // High resolution (e.g., 300 DPI) to retain visual fidelity.
            Resolution = 300,
            // Use high‑quality rendering algorithms.
            UseHighQualityRendering = true
        };

        // Configure metafile rendering to keep vector graphics as vectors.
        pngOptions.MetafileRenderingOptions.RenderingMode = MetafileRenderingMode.Vector;

        // Save the first page of the PDF as PNG.
        pdfDoc.Save(pngPath, pngOptions);

        // -----------------------------------------------------------------
        // Step 3: Validate that the PNG was created and contains data.
        // -----------------------------------------------------------------
        if (!File.Exists(pngPath) || new FileInfo(pngPath).Length == 0)
            throw new InvalidOperationException("PNG conversion failed or produced an empty file.");

        // Indicate successful completion.
        Console.WriteLine("PDF successfully converted to high‑resolution PNG:");
        Console.WriteLine(pngPath);
    }
}
