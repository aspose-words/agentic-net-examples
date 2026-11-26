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

        // Ensure the document has at least one paragraph to host the group shape.
        Paragraph paragraph = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(paragraph);

        // Create a GroupShape.
        GroupShape group = new GroupShape(doc)
        {
            Width = 400,
            Height = 300,
            Left = 0,
            Top = 0,
            WrapType = WrapType.None,
            RelativeHorizontalPosition = RelativeHorizontalPosition.Page,
            RelativeVerticalPosition = RelativeVerticalPosition.Page
        };

        // ---------- Create a sample picture ----------
        // Write a tiny PNG (1x1 transparent pixel) to a temporary file.
        string imagePath = Path.Combine(Path.GetTempPath(), "sample_image.png");
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        File.WriteAllBytes(imagePath, Convert.FromBase64String(base64Png));

        Shape picture = new Shape(doc, ShapeType.Image)
        {
            Width = 100,
            Height = 100,
            Left = 0,
            Top = 0
        };
        picture.ImageData.SetImage(imagePath);
        group.AppendChild(picture);

        // ---------- Create a textbox ----------
        Shape textBox = new Shape(doc, ShapeType.TextBox)
        {
            Width = 150,
            Height = 80,
            Left = 120,
            Top = 0,
            FillColor = System.Drawing.Color.LightYellow,
            StrokeColor = System.Drawing.Color.DarkGray
        };
        // Add text to the textbox.
        Paragraph tbParagraph = new Paragraph(doc);
        tbParagraph.AppendChild(new Run(doc, "Sample TextBox"));
        textBox.AppendChild(tbParagraph);
        group.AppendChild(textBox);

        // ---------- Create an AutoShape (Rectangle) ----------
        Shape autoShape = new Shape(doc, ShapeType.Rectangle)
        {
            Width = 120,
            Height = 100,
            Left = 0,
            Top = 120,
            FillColor = System.Drawing.Color.LightGreen,
            StrokeColor = System.Drawing.Color.DarkGreen
        };
        // Add a caption inside the auto shape.
        Paragraph autoParagraph = new Paragraph(doc);
        autoParagraph.AppendChild(new Run(doc, "AutoShape"));
        autoShape.AppendChild(autoParagraph);
        group.AppendChild(autoShape);

        // Insert the group shape into the document.
        paragraph.AppendChild(group);

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "GroupShapeExample.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");

        // Clean up temporary image file.
        if (File.Exists(imagePath))
            File.Delete(imagePath);
    }
}
