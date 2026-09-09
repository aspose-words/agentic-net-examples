using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample Word document with a vector shape.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        // Insert a rectangle shape (vector graphic).
        builder.InsertShape(ShapeType.Rectangle, 300, 150);
        // Save the document as PDF – this PDF will contain the vector graphic.
        const string pdfPath = "input.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the source PDF.");

        // Step 2: Load the PDF and render it to a high‑resolution PNG.
        Document pdfDoc = new Document(pdfPath);
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // High DPI to retain visual fidelity.
            Resolution = 300,
            // Ensure high‑quality rendering algorithms are used.
            UseHighQualityRendering = true,
            // Render the first page (index 0). Adjust if multiple pages are needed.
            PageSet = new PageSet(0)
        };

        const string pngPath = "output.png";
        pdfDoc.Save(pngPath, pngOptions);

        // Step 3: Validate that the PNG was created and contains data.
        if (!File.Exists(pngPath) || new FileInfo(pngPath).Length == 0)
            throw new InvalidOperationException("The PNG conversion failed or produced an empty file.");

        // Optional: Inform the user (no interactive input required).
        Console.WriteLine($"PDF successfully converted to high‑resolution PNG: {pngPath}");
    }
}
