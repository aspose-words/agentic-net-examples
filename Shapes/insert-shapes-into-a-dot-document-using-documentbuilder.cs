using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

namespace ShapeInsertionExample
{
    class Program
    {
        static void Main()
        {
            // Create a new empty document and associate a DocumentBuilder with it.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a paragraph of text before the shapes.
            builder.Writeln("Below are some shapes inserted using DocumentBuilder:");

            // Insert an inline rectangle shape (width: 100 points, height: 50 points).
            Shape inlineRectangle = builder.InsertShape(ShapeType.Rectangle, 100, 50);
            // Optional: set a fill color for the rectangle.
            inlineRectangle.Fill.ForeColor = System.Drawing.Color.LightBlue;
            // Move the cursor after the shape to continue inserting content.
            builder.Writeln();

            // Insert a floating ellipse shape positioned 100 points from the left and top of the page.
            Shape floatingEllipse = builder.InsertShape(
                ShapeType.Ellipse,
                RelativeHorizontalPosition.Page, 100,   // Horizontal position
                RelativeVerticalPosition.Page, 100,     // Vertical position
                80, 80,                                 // Width and height
                WrapType.None);                         // No text wrapping

            // Set additional properties for the floating shape.
            floatingEllipse.Fill.ForeColor = System.Drawing.Color.LightCoral;
            floatingEllipse.Stroke.Color = System.Drawing.Color.DarkRed;

            // Insert another paragraph after the floating shape.
            builder.Writeln();
            builder.Writeln("Shapes have been inserted.");

            // Save the document to a DOCX file.
            doc.Save("ShapesInserted.docx", SaveFormat.Docx);
        }
    }
}
