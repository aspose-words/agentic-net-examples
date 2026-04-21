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

        // Create a group shape with an initial size of 200x200 points.
        GroupShape group = new GroupShape(doc);
        group.Bounds = new RectangleF(0, 0, 200, 200);
        // Use a 1:1 coordinate system so that child shape coordinates are in points.
        group.CoordSize = new Size(200, 200);
        group.CoordOrigin = new Point(0, 0);

        // First child shape – a red rectangle positioned at (0,0) inside the group.
        Shape rect = new Shape(doc, ShapeType.Rectangle)
        {
            Width = 100,
            Height = 100,
            Left = 0,
            Top = 0,
            FillColor = Color.Red
        };
        group.AppendChild(rect);

        // Insert the group into the document.
        builder.InsertNode(group);

        // Save the intermediate document (optional, just to illustrate the state before appending).
        string intermediatePath = Path.Combine(Directory.GetCurrentDirectory(), "GroupShapeInitial.docx");
        doc.Save(intermediatePath);

        // -----------------------------------------------------------------
        // Append a new shape – a blue ellipse – to the existing group shape.
        // -----------------------------------------------------------------
        Shape ellipse = new Shape(doc, ShapeType.Ellipse)
        {
            Width = 150,
            Height = 150,
            // Position the ellipse so that it extends beyond the current group bounds.
            Left = 180,
            Top = 180,
            FillColor = Color.Blue
        };
        group.AppendChild(ellipse);

        // Update the group bounds to encompass the newly added shape.
        // Since we use a 1:1 coordinate system, child coordinates are directly comparable.
        float newRight = (float)Math.Max(group.Bounds.Right, ellipse.Left + ellipse.Width);
        float newBottom = (float)Math.Max(group.Bounds.Bottom, ellipse.Top + ellipse.Height);
        // Left and Top remain unchanged (0,0) in this example.
        group.Bounds = new RectangleF(group.Bounds.X, group.Bounds.Y, newRight - group.Bounds.X, newBottom - group.Bounds.Y);
        // Keep the coordinate system in sync with the new size.
        group.CoordSize = new Size((int)group.Bounds.Width, (int)group.Bounds.Height);

        // Save the final document.
        string finalPath = Path.Combine(Directory.GetCurrentDirectory(), "GroupShapeUpdated.docx");
        doc.Save(finalPath);

        // Validation – ensure the output file was created.
        if (!File.Exists(finalPath))
            throw new Exception("The document was not saved correctly.");
    }
}
