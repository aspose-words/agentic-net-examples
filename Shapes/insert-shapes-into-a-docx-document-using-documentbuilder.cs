using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class InsertShapesExample
{
    static void Main()
    {
        // Create a new empty document and associate a DocumentBuilder with it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // 1. Insert an inline shape (a simple rectangle) with a fixed size.
        // -----------------------------------------------------------------
        // The shape is placed directly in the text flow.
        Shape inlineShape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        // Optional: set a fill color and a line color.
        inlineShape.Fill.Color = Color.LightBlue;
        inlineShape.Stroke.Color = Color.DarkBlue;

        // Add a paragraph break after the inline shape.
        builder.Writeln();

        // -----------------------------------------------------------------
        // 2. Insert a floating shape (a rounded rectangle) with custom position.
        // -----------------------------------------------------------------
        // The shape is positioned relative to the page and does not affect text flow.
        Shape floatingShape = builder.InsertShape(
            ShapeType.TopCornersRounded,                     // shape type
            RelativeHorizontalPosition.Page, 150,           // left position (points)
            RelativeVerticalPosition.Page, 150,             // top position (points)
            120,                                            // width (points)
            80,                                             // height (points)
            WrapType.None);                                 // no text wrapping

        // Set additional formatting for the floating shape.
        floatingShape.Fill.Color = Color.LightGreen;
        floatingShape.Stroke.Color = Color.Green;
        floatingShape.BehindText = true; // place behind the text.

        // Add another paragraph break.
        builder.Writeln();

        // -----------------------------------------------------------------
        // 3. Group two shapes together.
        // -----------------------------------------------------------------
        // Create two individual shapes that will be grouped.
        Shape shape1 = builder.InsertShape(ShapeType.Ellipse, 80, 80);
        shape1.Fill.Color = Color.Pink;
        shape1.Stroke.Color = Color.Red;

        Shape shape2 = builder.InsertShape(ShapeType.Star, 80, 80);
        shape2.Fill.Color = Color.Yellow;
        shape2.Stroke.Color = Color.Orange;

        // Group the two shapes. The group will be inserted at the current cursor position.
        GroupShape group = builder.InsertGroupShape(shape1, shape2);
        // Optionally adjust the group's position.
        group.Left = 300;
        group.Top = 200;

        // -----------------------------------------------------------------
        // 4. Save the document with OOXML compliance that supports DML shapes.
        // -----------------------------------------------------------------
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Transitional
        };

        doc.Save("ShapesInsertion.docx", saveOptions);
    }
}
