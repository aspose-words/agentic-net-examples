using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an empty paragraph to host the group shape.
        Paragraph para = builder.CurrentParagraph;

        // Create a GroupShape and add it to the document.
        GroupShape group = new GroupShape(doc);
        para.AppendChild(group);

        // Create an initial rectangle shape inside the group.
        Shape rect1 = new Shape(doc, ShapeType.Rectangle);
        rect1.Width = 100;
        rect1.Height = 50;
        rect1.Left = 0;
        rect1.Top = 0;
        rect1.StrokeColor = Color.Black;
        rect1.FillColor = Color.LightBlue;
        group.AppendChild(rect1);

        // Append a new shape (ellipse) to the existing GroupShape.
        Shape ellipse = new Shape(doc, ShapeType.Ellipse);
        ellipse.Width = 80;
        ellipse.Height = 80;
        // Position the new shape within the group.
        ellipse.Left = 120;
        ellipse.Top = 20;
        ellipse.StrokeColor = Color.DarkGreen;
        ellipse.FillColor = Color.LightGreen;
        group.AppendChild(ellipse);

        // Recalculate the bounds of the group to encompass all child shapes.
        RectangleF newBounds = CalculateGroupBounds(group);

        // Update the group's size and position based on the calculated bounds.
        group.Width = newBounds.Width;
        group.Height = newBounds.Height;
        group.Left = newBounds.Left;
        group.Top = newBounds.Top;

        // Save the document.
        string outputPath = "GroupShapeExample.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to create the output file: {outputPath}");
    }

    // Helper method to compute the union of bounds of all child shapes in a group.
    private static RectangleF CalculateGroupBounds(GroupShape group)
    {
        RectangleF bounds = RectangleF.Empty;
        foreach (Shape child in group.GetChildNodes(NodeType.Shape, false))
        {
            RectangleF childBounds = child.Bounds;
            if (bounds.IsEmpty)
                bounds = childBounds;
            else
                bounds = RectangleF.Union(bounds, childBounds);
        }
        return bounds;
    }
}
