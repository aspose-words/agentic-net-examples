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

        // Insert two floating shapes that will be grouped.
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 200, 150);
        rectangle.Left = 50;               // Position from the left edge of the page.
        rectangle.Top = 50;                // Position from the top edge of the page.
        rectangle.Stroke.Color = Color.Blue;

        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 150);
        ellipse.Left = 120;
        ellipse.Top = 80;
        ellipse.Stroke.Color = Color.Green;

        // Group the shapes. The InsertGroupShape method automatically calculates the
        // position and size of the new GroupShape and inserts it at the current cursor location.
        GroupShape group = builder.InsertGroupShape(rectangle, ellipse);

        // Example of setting additional properties on the group shape.
        group.WrapType = WrapType.None;    // Make the group floating.
        group.BehindText = true;           // Place it behind the document text.

        // Save the document as an RTF file.
        doc.Save("GroupShape.rtf", SaveFormat.Rtf);
    }
}
