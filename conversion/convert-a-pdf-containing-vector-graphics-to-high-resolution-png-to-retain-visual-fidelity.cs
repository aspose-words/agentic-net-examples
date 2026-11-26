using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample document containing a vector shape (rectangle).
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        // Insert a rectangle shape (vector graphic) with specific dimensions.
        builder.InsertShape(ShapeType.Rectangle, 200, 100);

        // Save the document as PDF – the PDF will contain the vector shape.
        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The sample PDF was not created.");

        // Step 2: Load the generated PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Step 3: Configure high‑resolution PNG export options.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Set a high resolution (e.g., 300 DPI) to retain visual fidelity.
            Resolution = 300,
            // Export only the first page (adjust as needed for multi‑page PDFs).
            PageSet = new PageSet(0)
        };

        // Step 4: Save the PDF as a high‑resolution PNG image.
        const string pngPath = "output.png";
        pdfDoc.Save(pngPath, pngOptions);

        // Step 5: Validate that the PNG file was created.
        if (!File.Exists(pngPath))
            throw new InvalidOperationException("The PNG image was not created.");

        // Indicate success.
        Console.WriteLine("PDF successfully converted to high‑resolution PNG.");
    }
}
