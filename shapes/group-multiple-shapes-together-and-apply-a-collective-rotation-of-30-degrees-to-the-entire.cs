using System;
using System.IO;
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

        // Ensure there is a paragraph to host the group shape.
        Paragraph paragraph = builder.CurrentParagraph;

        // Create a GroupShape and set its size and position.
        GroupShape group = new GroupShape(doc)
        {
            Width = 300,
            Height = 200,
            Left = 100,
            Top = 100,
            WrapType = WrapType.None
        };

        // First shape: a rectangle.
        Shape rect = new Shape(doc, ShapeType.Rectangle)
        {
            Width = 120,
            Height = 80,
            Left = 0,
            Top = 0,
            StrokeColor = Color.Black,
            FillColor = Color.LightBlue
        };
        group.AppendChild(rect);

        // Second shape: an ellipse.
        Shape ellipse = new Shape(doc, ShapeType.Ellipse)
        {
            Width = 100,
            Height = 100,
            Left = 150,
            Top = 50,
            StrokeColor = Color.DarkGreen,
            FillColor = Color.LightGreen
        };
        group.AppendChild(ellipse);

        // Apply a collective rotation of 30 degrees to the entire group.
        group.Rotation = 30;

        // Insert the group shape into the document.
        paragraph.AppendChild(group);

        // Save the document.
        string outputPath = "GroupedShapes.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }
}
