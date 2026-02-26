using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

namespace GroupShapeToPdfExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to add content to the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert two individual shapes.
            Shape rect = builder.InsertShape(ShapeType.Rectangle, 200, 150);
            rect.Left = 50;
            rect.Top = 50;
            rect.Stroke.Color = Color.Red;

            Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 150);
            ellipse.Left = 300;
            ellipse.Top = 80;
            ellipse.Stroke.Color = Color.Blue;

            // Group the two shapes into a GroupShape and insert it at the current position.
            GroupShape group = builder.InsertGroupShape(rect, ellipse);

            // Optionally adjust group properties (e.g., position).
            group.Left = 20;
            group.Top = 20;

            // Save the document as PDF.
            // Using PdfSaveOptions to ensure proper rendering of drawing shapes.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Render DrawingML shapes directly.
                DmlRenderingMode = DmlRenderingMode.DrawingML
            };

            doc.Save("GroupShapeDocument.pdf", pdfOptions);
        }
    }
}
