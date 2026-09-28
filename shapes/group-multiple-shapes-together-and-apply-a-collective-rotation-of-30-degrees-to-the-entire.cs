using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class GroupShapeRotationExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a paragraph to host the group shape (Body can contain only block-level nodes).
        Paragraph paragraph = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(paragraph);

        // Create a GroupShape and set its size and position.
        GroupShape group = new GroupShape(doc)
        {
            Width = 300,          // Width of the group.
            Height = 200,         // Height of the group.
            Left = 100,           // Horizontal position on the page.
            Top = 100,            // Vertical position on the page.
            WrapType = WrapType.None // No text wrapping.
        };
        // Insert the group into the paragraph.
        paragraph.AppendChild(group);

        // First shape: a rectangle.
        Shape rect = new Shape(doc, ShapeType.Rectangle)
        {
            Width = 100,
            Height = 50,
            Left = 0,               // Position relative to the group.
            Top = 0,
            FillColor = Color.LightBlue,
            WrapType = WrapType.None
        };
        group.AppendChild(rect);

        // Second shape: an ellipse.
        Shape ellipse = new Shape(doc, ShapeType.Ellipse)
        {
            Width = 80,
            Height = 80,
            Left = 120,          // Position relative to the group.
            Top = 30,
            FillColor = Color.LightGreen,
            WrapType = WrapType.None
        };
        group.AppendChild(ellipse);

        // Apply a collective rotation of 30 degrees to the entire group.
        group.Rotation = 30f;

        // Save the document.
        string outputPath = "GroupShapeRotation.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }
}
