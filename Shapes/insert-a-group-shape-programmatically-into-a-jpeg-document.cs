using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class InsertGroupShapeIntoJpeg
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a group shape that will contain other shapes.
        GroupShape group = new GroupShape(doc);
        // Set the size and position of the group shape (in points).
        group.Bounds = new RectangleF(0, 0, 200, 200);
        // Make the group floating and place it at the center of the page.
        group.WrapType = WrapType.None;
        group.BehindText = true;
        group.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        group.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        group.HorizontalAlignment = HorizontalAlignment.Center;
        group.VerticalAlignment = VerticalAlignment.Center;

        // Append the group shape to the current paragraph.
        builder.CurrentParagraph.AppendChild(group);

        // Create a rectangle shape to be a child of the group.
        Shape rect = new Shape(doc, ShapeType.Rectangle);
        rect.Width = 100;
        rect.Height = 100;
        rect.Left = 0;
        rect.Top = 0;
        rect.FillColor = Color.LightBlue;
        // Add the rectangle to the group.
        group.AppendChild(rect);

        // Save the document as a JPEG image.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Jpeg);
        doc.Save("GroupShapeOutput.jpeg", saveOptions);
    }
}
