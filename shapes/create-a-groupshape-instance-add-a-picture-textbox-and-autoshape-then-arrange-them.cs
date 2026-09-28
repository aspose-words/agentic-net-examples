using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a temporary PNG image (1x1 pixel) from a Base64 string.
        string imagePath = Path.Combine(Path.GetTempPath(), "sample_image.png");
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");
        File.WriteAllBytes(imagePath, pngBytes);

        // Create a GroupShape.
        GroupShape group = new GroupShape(doc)
        {
            Width = 300,
            Height = 300,
            Left = 100,
            Top = 100,
            WrapType = WrapType.None
        };

        // ---- Picture shape ----
        Shape picture = new Shape(doc, ShapeType.Image)
        {
            Width = 100,
            Height = 100,
            Left = 0,
            Top = 0
        };
        picture.ImageData.SetImage(imagePath);
        group.AppendChild(picture);

        // ---- TextBox shape ----
        Shape textBox = new Shape(doc, ShapeType.TextBox)
        {
            Width = 150,
            Height = 80,
            Left = 110,
            Top = 0,
            FillColor = System.Drawing.Color.LightYellow,
            StrokeColor = System.Drawing.Color.DarkGray
        };
        // Add text to the textbox.
        Paragraph tbParagraph = new Paragraph(doc);
        Run tbRun = new Run(doc, "Sample TextBox");
        tbParagraph.AppendChild(tbRun);
        textBox.AppendChild(tbParagraph);
        group.AppendChild(textBox);

        // ---- AutoShape (Rectangle) ----
        Shape autoShape = new Shape(doc, ShapeType.Rectangle)
        {
            Width = 120,
            Height = 120,
            Left = 0,
            Top = 110,
            FillColor = System.Drawing.Color.LightGreen,
            StrokeColor = System.Drawing.Color.DarkGreen
        };
        // Add text to the auto shape.
        Paragraph asParagraph = new Paragraph(doc);
        Run asRun = new Run(doc, "AutoShape");
        asParagraph.AppendChild(asRun);
        autoShape.AppendChild(asParagraph);
        group.AppendChild(autoShape);

        // Insert the group shape into the document.
        builder.InsertNode(group);

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "GroupShapeExample.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");

        // Clean up temporary image.
        if (File.Exists(imagePath))
            File.Delete(imagePath);
    }
}
