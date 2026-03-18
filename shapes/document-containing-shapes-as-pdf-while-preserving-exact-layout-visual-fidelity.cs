using System;
using System.Drawing;
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
        rectangle.FillColor = Color.LightBlue;

        // Insert an inline picture shape.
        Shape picture = builder.InsertImage(@"C:\Images\Sample.png");
        picture.Width = 150;
        picture.Height = 150;

        // Configure PDF save options to preserve exact layout and visual fidelity.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Render DrawingML shapes directly (no fallback).
            DmlRenderingMode = DmlRenderingMode.DrawingML,
            // Use high‑quality (slow) rendering algorithms.
            UseHighQualityRendering = true,
            // Enable anti‑aliasing for smoother edges.
            UseAntiAliasing = true,
            // Cache background graphics to keep layout consistent.
            CacheBackgroundGraphics = true
        };

        // Save the document as PDF using the configured options.
        doc.Save(@"C:\Output\ShapesDocument.pdf", pdfOptions);
    }
}
