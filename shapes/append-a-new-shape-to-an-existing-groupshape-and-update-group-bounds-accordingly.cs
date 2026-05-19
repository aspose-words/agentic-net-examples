using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a group shape and set its initial bounds.
        GroupShape group = new GroupShape(doc);
        group.Bounds = new RectangleF(0, 0, 300, 300); // Initial size 300x300 points.

        // First child shape – a rectangle.
        Shape rect = new Shape(doc, ShapeType.Rectangle)
        {
            Width = 100,
            Height = 80,
            Left = 20,
            Top = 20,
            FillColor = Color.LightBlue,
            Stroke = { Color = Color.DarkBlue }
        };
        group.AppendChild(rect);

        // Second child shape – an ellipse.
        Shape ellipse = new Shape(doc, ShapeType.Ellipse)
        {
            Width = 120,
            Height = 90,
            Left = 150,
            Top = 100,
            FillColor = Color.LightGreen,
            Stroke = { Color = Color.DarkGreen }
        };
        group.AppendChild(ellipse);

        // Insert the group shape into the document.
        builder.InsertNode(group);

        // Append a new shape – a star – to the existing group.
        Shape star = new Shape(doc, ShapeType.Star)
        {
            Width = 80,
            Height = 80,
            Left = 250,   // Position that extends beyond the original bounds.
            Top = 250,
            FillColor = Color.Yellow,
            Stroke = { Color = Color.Orange }
        };
        group.AppendChild(star);

        // Update the group bounds to encompass the new shape.
        // Expand the bounds manually to cover the furthest extents.
        float newRight = Math.Max(group.Bounds.Right, (float)(star.Left + star.Width));
        float newBottom = Math.Max(group.Bounds.Bottom, (float)(star.Top + star.Height));
        group.Bounds = new RectangleF(group.Bounds.X, group.Bounds.Y, newRight - group.Bounds.X, newBottom - group.Bounds.Y);

        // Validation: ensure the group now contains three child shapes.
        if (group.Count != 3)
            throw new InvalidOperationException("The group shape does not contain the expected number of child shapes.");

        // Validation: ensure the updated bounds are large enough.
        if (group.Bounds.Width < 330 || group.Bounds.Height < 330)
            throw new InvalidOperationException("The group bounds were not updated correctly.");

        // Save the document.
        string outputPath = "AppendShapeToGroup.docx";
        doc.Save(outputPath);
    }
}
