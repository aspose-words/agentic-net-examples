using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create an initial GroupShape with a defined bounding rectangle.
        GroupShape group = new GroupShape(doc);
        group.Bounds = new RectangleF(0, 0, 200, 200); // Initial size 200x200 points.

        // Add a rectangle shape inside the group.
        Shape rect = new Shape(doc, ShapeType.Rectangle)
        {
            Width = 100,
            Height = 100,
            Left = 20,
            Top = 20,
            FillColor = Color.LightBlue,
            Stroke = { Color = Color.DarkBlue }
        };
        group.AppendChild(rect);

        // Insert the group into the document.
        builder.InsertNode(group);

        // Create a new shape that will be appended to the existing group.
        Shape ellipse = new Shape(doc, ShapeType.Ellipse)
        {
            Width = 120,
            Height = 80,
            Left = 150,   // Position that extends beyond the original group bounds.
            Top = 150,
            FillColor = Color.LightCoral,
            Stroke = { Color = Color.Maroon }
        };

        // Append the new shape to the group.
        group.AppendChild(ellipse);

        // Update the group's bounds to encompass all child shapes.
        // Use RectangleF.Union to combine the existing bounds with the new shape's bounds.
        group.Bounds = RectangleF.Union(group.Bounds, ellipse.Bounds);

        // Validate that the bounds have been updated correctly.
        // Expected width >= 270 (200 original + extra 70) and height >= 230.
        if (group.Bounds.Width < 270 || group.Bounds.Height < 230)
            throw new InvalidOperationException("Group bounds were not updated correctly.");

        // Save the document.
        string outputPath = "GroupShapeAppend.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
