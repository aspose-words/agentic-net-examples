using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class InsertGroupShapeIntoRtf
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Ensure there is a paragraph to host the group shape.
        builder.Writeln("Paragraph before the group shape.");

        // Create a group shape. The constructor requires a DocumentBase (Document).
        GroupShape group = new GroupShape(doc);

        // Set the size and position of the group shape (in points).
        group.Width = 200;
        group.Height = 100;
        group.Left = 50;
        group.Top = 50;

        // Set wrapping to none so the group behaves as a floating object.
        group.WrapType = WrapType.None;
        group.BehindText = true;
        group.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        group.RelativeVerticalPosition = RelativeVerticalPosition.Page;

        // Create a rectangle shape to add to the group.
        Shape rect = new Shape(doc, ShapeType.Rectangle);
        rect.Width = 80;
        rect.Height = 60;
        rect.Left = 0;   // Position relative to the group's coordinate space.
        rect.Top = 0;
        rect.Fill.Color = System.Drawing.Color.LightBlue;
        rect.StrokeColor = System.Drawing.Color.DarkBlue;
        rect.StrokeWeight = 0.5;

        // Add the rectangle as a child of the group shape.
        group.AppendChild(rect);

        // Create a second shape (e.g., an ellipse) inside the same group.
        Shape ellipse = new Shape(doc, ShapeType.Ellipse);
        ellipse.Width = 80;
        ellipse.Height = 60;
        ellipse.Left = 100; // Position next to the rectangle within the group.
        ellipse.Top = 0;
        ellipse.Fill.Color = System.Drawing.Color.LightCoral;
        ellipse.StrokeColor = System.Drawing.Color.DarkRed;
        ellipse.StrokeWeight = 0.5;

        group.AppendChild(ellipse);

        // Insert the group shape into the document after the current paragraph.
        // The group shape must be added to the document's node collection.
        builder.CurrentParagraph.AppendChild(group);

        // Add another paragraph after the group shape for clarity.
        builder.Writeln("\nParagraph after the group shape.");

        // Save the document as RTF using RtfSaveOptions.
        RtfSaveOptions saveOptions = new RtfSaveOptions
        {
            // Ensure the document is saved in RTF format.
            SaveFormat = SaveFormat.Rtf
        };
        doc.Save("GroupShapeOutput.rtf", saveOptions);
    }
}
