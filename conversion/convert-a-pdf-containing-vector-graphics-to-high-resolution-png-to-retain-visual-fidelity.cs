using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary PDF and the final PNG.
        const string inputPdfPath = "input.pdf";
        const string outputPngPath = "output.png";

        // -----------------------------------------------------------------
        // Step 1: Create a sample PDF that contains vector graphics.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Add some text.
        builder.Writeln("Sample PDF with vector graphics:");

        // Insert a vector shape (a 5‑point star). Use a shape type that exists in the API.
        // ShapeType.Star5 is not available; use ShapeType.Star5 (fallback to a regular star shape).
        builder.InsertShape(ShapeType.Star, 200, 200);

        // Save the document as PDF – this PDF now contains vector graphics.
        sourceDoc.Save(inputPdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Step 2: Load the PDF and convert the first page to a high‑resolution PNG.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(inputPdfPath);

        // Configure image save options for high quality.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Render at 300 DPI to retain visual fidelity.
            Resolution = 300,
            // Use the high‑quality rendering algorithm.
            UseHighQualityRendering = true
        };

        // Save the first page of the PDF as a PNG image.
        pdfDoc.Save(outputPngPath, pngOptions);

        // -----------------------------------------------------------------
        // Step 3: Verify that the PNG file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPngPath))
            throw new InvalidOperationException("Expected output PNG was not created.");
    }
}
