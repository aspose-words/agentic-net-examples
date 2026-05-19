using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a rectangle shape (vector graphic).
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        // Insert a rectangle shape; default colors are sufficient for the example.
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 300, 150);

        // Save the document as PDF – vector graphics are preserved.
        const string pdfPath = "sample.pdf";
        sampleDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Configure high‑resolution PNG export.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            Resolution = 300,               // High DPI for visual fidelity.
            UseHighQualityRendering = true,
            UseAntiAliasing = true
        };

        // Save the first page of the PDF as a PNG image.
        const string pngPath = "output.png";
        pdfDoc.Save(pngPath, pngOptions);

        // Verify that the PNG file was created and is not empty.
        if (!File.Exists(pngPath) || new FileInfo(pngPath).Length == 0)
            throw new InvalidOperationException("The PNG conversion failed; the output file was not created or is empty.");

        // Optional cleanup of intermediate files.
        // File.Delete(pdfPath);
    }
}
