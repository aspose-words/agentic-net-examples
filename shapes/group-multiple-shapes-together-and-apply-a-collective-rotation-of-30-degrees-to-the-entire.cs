using System;
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

        // Ensure there is a paragraph to host the group shape.
        Paragraph para = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(para);

        // Create a group shape.
        GroupShape group = new GroupShape(doc)
        {
            // Set size and position of the group.
            Width = 200,
            Height = 200,
            Left = 100,
            Top = 100,
            // Apply a collective rotation of 30 degrees.
            Rotation = 30
        };

        // First shape: a rectangle.
        Shape rect = new Shape(doc, ShapeType.Rectangle)
        {
            Width = 80,
            Height = 80,
            FillColor = System.Drawing.Color.LightBlue,
            StrokeColor = System.Drawing.Color.Blue,
            // Position relative to the group.
            Left = 0,
            Top = 0
        };
        group.AppendChild(rect);

        // Second shape: an ellipse.
        Shape ellipse = new Shape(doc, ShapeType.Ellipse)
        {
            Width = 80,
            Height = 80,
            FillColor = System.Drawing.Color.LightCoral,
            StrokeColor = System.Drawing.Color.Red,
            // Position relative to the group.
            Left = 120,
            Top = 0
        };
        group.AppendChild(ellipse);

        // Insert the group shape into the document.
        para.AppendChild(group);

        // Save the document.
        string outputPath = "GroupedShapes.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }
}
