using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating rectangle shape.
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 200, 100);
        rectangle.WrapType = WrapType.None; // Float without text wrapping.
        rectangle.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        rectangle.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        rectangle.Left = 100; // Position from the left edge of the page.
        rectangle.Top = 150;  // Position from the top edge of the page.
        rectangle.FillColor = System.Drawing.Color.LightBlue;

        // Insert a simple in‑memory PNG image (a 16×16 red square).
        const string base64Png = 
            "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAIAAACQkWg2AAAADklEQVR4nGNgGAWjYBSMABcAAf8B9Z0AAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        using (MemoryStream ms = new MemoryStream(pngBytes))
        {
            Shape picture = builder.InsertImage(ms);
            picture.Width = 150;
            picture.Height = 150;
        }

        // Configure PDF save options to preserve exact layout and visual fidelity.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            DmlRenderingMode = DmlRenderingMode.DrawingML,
            UseHighQualityRendering = true,
            UseAntiAliasing = true,
            CacheBackgroundGraphics = true
        };

        // Save the document as PDF in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ShapesDocument.pdf");
        doc.Save(outputPath, pdfOptions);

        Console.WriteLine($"PDF saved to: {outputPath}");
    }
}
