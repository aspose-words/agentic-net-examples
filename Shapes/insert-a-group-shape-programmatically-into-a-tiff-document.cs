using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class InsertGroupShapeIntoTiff
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Get the current paragraph (the builder is positioned at the start of the document).
        Paragraph paragraph = builder.CurrentParagraph;

        // Create a group shape and set its size and position.
        GroupShape group = new GroupShape(doc);
        // Bounds are defined in points. Here we create a 200x200 points group at (0,0).
        group.Bounds = new RectangleF(0, 0, 200, 200);
        // The coordinate space for child shapes uses integer types (Point and Size).
        group.CoordOrigin = new Point(0, 0);
        group.CoordSize   = new Size(200, 200);

        // Add the group shape to the document.
        paragraph.AppendChild(group);

        // Create a rectangle shape to place inside the group.
        Shape innerShape = new Shape(doc, ShapeType.Rectangle);
        innerShape.Width = 100;
        innerShape.Height = 100;
        innerShape.Left = 50;   // Position within the group's coordinate space.
        innerShape.Top = 50;
        innerShape.Fill.Color = Color.LightBlue;
        innerShape.Stroke.Color = Color.DarkBlue;

        // Add the rectangle to the group.
        group.AppendChild(innerShape);

        // Save the document as a TIFF image.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Example: use LZW compression (optional).
            TiffCompression = TiffCompression.Lzw
        };

        doc.Save("GroupShapeOutput.tiff", tiffOptions);
    }
}
