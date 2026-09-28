using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class GroupShapeExample
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Ensure there is a paragraph to host the group shape.
        builder.Writeln();

        // Create a GroupShape with initial bounds.
        GroupShape group = new GroupShape(doc);
        group.Bounds = new RectangleF(0, 0, 200, 200);

        // Append the group shape to the current paragraph.
        builder.CurrentParagraph.AppendChild(group);

        // Create a rectangle shape to add to the group.
        Shape rect = new Shape(doc, ShapeType.Rectangle);
        rect.Width = 100;
        rect.Height = 50;
        rect.Left = 20;   // Position relative to the group.
        rect.Top = 30;
        rect.StrokeColor = Color.Blue;
        rect.FillColor = Color.LightGray;

        // Append the rectangle shape to the group.
        group.AppendChild(rect);

        // Update the group bounds to encompass the new shape.
        // Here we simply expand the bounds to ensure the rectangle fits.
        group.Bounds = new RectangleF(0, 0, 250, 250);

        // Save the document.
        string outputPath = "GroupShapeExample.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }
}
