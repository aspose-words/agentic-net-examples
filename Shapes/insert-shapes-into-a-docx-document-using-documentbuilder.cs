using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class InsertShapesExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // 1. Insert an inline rectangle shape (auto shape) with a red border.
        // -----------------------------------------------------------------
        // InsertShape inserts a shape of the specified type. The parameters after the
        // ShapeType are width and height in points for an inline shape.
        Shape inlineRect = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        // Set the line color and thickness.
        inlineRect.StrokeColor = Color.Red;
        inlineRect.StrokeWeight = 2.0;
        // Add a caption below the shape.
        builder.Writeln();
        builder.Writeln("Inline rectangle shape above.");

        // -----------------------------------------------------------------
        // 2. Insert a floating ellipse shape positioned at (100,100) points on the page.
        // -----------------------------------------------------------------
        // For a floating shape we need to specify its position relative to the page.
        Shape floatingEllipse = builder.InsertShape(
            ShapeType.Ellipse,                     // Shape type.
            RelativeHorizontalPosition.Page,       // Horizontal reference.
            100,                                   // Horizontal offset.
            RelativeVerticalPosition.Page,         // Vertical reference.
            100,                                   // Vertical offset.
            150,                                   // Width.
            100,                                   // Height.
            WrapType.None);                        // No text wrapping (shape behind text).
        // Set fill color and make the shape semi‑transparent.
        floatingEllipse.Fill.ForeColor = Color.LightBlue;
        floatingEllipse.Fill.Transparency = 0.3;
        // Place the shape behind the text.
        floatingEllipse.BehindText = true;

        // -----------------------------------------------------------------
        // 3. Insert an image as a watermark in the header (floating, centered).
        // -----------------------------------------------------------------
        // Move the cursor to the primary header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        // Insert the image.
        Shape watermark = builder.InsertImage("Logo.jpg");
        // Configure the image to behave as a watermark.
        watermark.WrapType = WrapType.None;
        watermark.BehindText = true;
        watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        watermark.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        // Center the image on the page.
        watermark.Left = (builder.PageSetup.PageWidth - watermark.Width) / 2;
        watermark.Top = (builder.PageSetup.PageHeight - watermark.Height) / 2;

        // -----------------------------------------------------------------
        // Save the document with ISO 29500:2008 Transitional compliance so that
        // non‑primitive shapes are stored using DML.
        // -----------------------------------------------------------------
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Transitional
        };
        doc.Save("InsertedShapes.docx", saveOptions);
    }
}
