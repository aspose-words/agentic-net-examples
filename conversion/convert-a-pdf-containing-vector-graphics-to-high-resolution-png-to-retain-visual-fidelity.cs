using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class PdfToPngConverter
{
    public static void Main()
    {
        // Step 1: Create a sample PDF with a vector rectangle shape.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        // Insert a rectangle shape (vector graphic) of size 200x100 points.
        builder.InsertShape(ShapeType.Rectangle, 200, 100);
        // Save the document as PDF.
        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The sample PDF was not created.");

        // Step 2: Load the PDF and convert the first page to a high‑resolution PNG.
        Document pdfDoc = new Document(pdfPath);
        const string pngPath = "output.png";

        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Set a high resolution (e.g., 300 DPI) to retain visual fidelity.
            Resolution = 300,
            // Export only the first page (index 0). To export all pages, omit this property.
            PageSet = new PageSet(0)
        };

        pdfDoc.Save(pngPath, pngOptions);

        // Verify that the PNG was created and is not empty.
        if (!File.Exists(pngPath) || new FileInfo(pngPath).Length == 0)
            throw new InvalidOperationException("The PNG image was not created or is empty.");

        Console.WriteLine("PDF successfully converted to high‑resolution PNG.");
    }
}
