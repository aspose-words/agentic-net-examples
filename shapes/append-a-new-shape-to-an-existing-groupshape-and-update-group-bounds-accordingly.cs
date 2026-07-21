using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class AppendShapeToGroup
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a group shape.
        GroupShape group = new GroupShape(doc);

        // First child shape – a rectangle.
        Shape rect = new Shape(doc, ShapeType.Rectangle)
        {
            Width = 100,
            Height = 50,
            Left = 0,
            Top = 0,
            Stroke = { Color = Color.Blue }
        };
        group.AppendChild(rect);

        // Second child shape – an ellipse.
        Shape ellipse = new Shape(doc, ShapeType.Ellipse)
        {
            Width = 80,
            Height = 80,
            Left = 120,
            Top = 20,
            Stroke = { Color = Color.Green }
        };
        group.AppendChild(ellipse);

        // Set initial bounds of the group to enclose the two shapes.
        group.Bounds = CalculateGroupBounds(group);

        // Insert the group shape into the document.
        builder.InsertNode(group);

        // --- Append a new shape to the existing group ---

        // New child shape – a star.
        Shape star = new Shape(doc, ShapeType.Star)
        {
            Width = 60,
            Height = 60,
            Left = 50,
            Top = 100,
            FillColor = Color.Yellow,
            Stroke = { Color = Color.Orange }
        };
        group.AppendChild(star);

        // Update the group's bounds to include the newly added shape.
        group.Bounds = CalculateGroupBounds(group);

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AppendShapeToGroup.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }

    // Helper method to compute the minimal bounding rectangle that contains all child shapes of a group.
    private static RectangleF CalculateGroupBounds(GroupShape group)
    {
        if (group.Count == 0)
            return new RectangleF(0, 0, 0, 0);

        float minLeft = float.MaxValue;
        float minTop = float.MaxValue;
        float maxRight = float.MinValue;
        float maxBottom = float.MinValue;

        foreach (ShapeBase child in group)
        {
            // Use the child's Bounds property (in the group's coordinate space).
            RectangleF childBounds = child.Bounds;

            if (childBounds.Left < minLeft) minLeft = childBounds.Left;
            if (childBounds.Top < minTop) minTop = childBounds.Top;
            if (childBounds.Right > maxRight) maxRight = childBounds.Right;
            if (childBounds.Bottom > maxBottom) maxBottom = childBounds.Bottom;
        }

        return new RectangleF(minLeft, minTop, maxRight - minLeft, maxBottom - minTop);
    }
}
