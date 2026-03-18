using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Check whether the document contains any revisions.
        if (doc.HasRevisions)
        {
            // Create a DocumentBuilder to modify the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the builder to the primary header (watermark will appear on every page).
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            // Create a shape that will serve as the watermark.
            Shape watermark = new Shape(doc, ShapeType.TextPlainText);
            watermark.TextPath.Text = "PENDING CHANGES";
            watermark.TextPath.FontFamily = "Arial";
            watermark.Width = 500;
            watermark.Height = 100;
            watermark.Rotation = -40; // Diagonal orientation.
            watermark.Fill.Color = Color.LightGray;
            watermark.StrokeColor = Color.LightGray;

            // Position the watermark relative to the page.
            watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            watermark.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            watermark.WrapType = WrapType.None;
            watermark.VerticalAlignment = VerticalAlignment.Center;
            watermark.HorizontalAlignment = HorizontalAlignment.Center;

            // Insert the watermark into the header.
            builder.InsertNode(watermark);
        }

        // Save the document (PDF format preserves the watermark).
        doc.Save("Output.pdf");
    }
}
