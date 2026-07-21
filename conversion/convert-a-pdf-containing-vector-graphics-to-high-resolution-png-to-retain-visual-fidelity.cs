using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary PDF and the resulting PNG.
        const string pdfPath = "input.pdf";
        const string pngPath = "output.png";

        // -----------------------------------------------------------------
        // Step 1: Create a sample Word document with vector graphics.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape (vector graphic) into the document.
        // Width and height are specified in points (1 point = 1/72 inch).
        builder.InsertShape(ShapeType.Rectangle, 200, 100);
        builder.Writeln(); // Add a line break after the shape.

        // Save the document as a PDF – this PDF now contains vector graphics.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Step 2: Load the PDF and convert it to a high‑resolution PNG.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // Configure image save options for PNG with high DPI (e.g., 300).
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            Resolution = 300 // Dots per inch for both horizontal and vertical resolution.
        };

        // Save the first page of the PDF as a PNG image.
        // (Aspose.Words renders only the first page when saving to an image format.)
        pdfDoc.Save(pngPath, pngOptions);

        // -----------------------------------------------------------------
        // Step 3: Validate that the PNG file was created successfully.
        // -----------------------------------------------------------------
        if (!File.Exists(pngPath))
            throw new InvalidOperationException("The PNG output file was not created.");

        FileInfo pngInfo = new FileInfo(pngPath);
        if (pngInfo.Length == 0)
            throw new InvalidOperationException("The PNG output file is empty.");

        // Optional: Inform the user (no interactive input required).
        Console.WriteLine($"Conversion completed. PNG saved to '{pngPath}' ({pngInfo.Length} bytes).");
    }
}
