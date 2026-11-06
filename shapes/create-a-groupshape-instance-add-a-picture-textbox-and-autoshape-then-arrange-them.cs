using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a GroupShape.
        GroupShape groupShape = new GroupShape(doc)
        {
            Width = 400,
            Height = 300,
            WrapType = WrapType.Inline
        };

        // ---------- Picture ----------
        // A minimal 1x1 PNG (transparent) encoded in base64.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X6V8AAAAASUVORK5CYII=");
        using (MemoryStream imageStream = new MemoryStream(pngBytes))
        {
            Shape pictureShape = new Shape(doc, ShapeType.Image)
            {
                Width = 100,
                Height = 100,
                Left = 0,
                Top = 0
            };
            pictureShape.ImageData.SetImage(imageStream);
            groupShape.AppendChild(pictureShape);
        }

        // ---------- TextBox ----------
        Shape textBoxShape = new Shape(doc, ShapeType.TextBox)
        {
            Width = 150,
            Height = 80,
            Left = 110,
            Top = 0
        };

        Paragraph tbParagraph = new Paragraph(doc);
        Run tbRun = new Run(doc, "Hello TextBox");
        tbParagraph.AppendChild(tbRun);
        textBoxShape.AppendChild(tbParagraph);
        groupShape.AppendChild(textBoxShape);

        // ---------- AutoShape (Rectangle) ----------
        Shape autoShape = new Shape(doc, ShapeType.Rectangle)
        {
            Width = 120,
            Height = 60,
            Left = 0,
            Top = 110,
            FillColor = System.Drawing.Color.Yellow,
            StrokeColor = System.Drawing.Color.Black
        };
        groupShape.AppendChild(autoShape);

        // Insert the group shape into the document.
        builder.InsertNode(groupShape);

        // Save the document.
        string outputPath = "GroupShapeExample.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to create the output file: {outputPath}");
    }
}
