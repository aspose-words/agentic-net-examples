using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class GroupShapeExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Insert a picture ----------
        // Use a tiny 1x1 PNG image (red pixel) encoded in base64.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ/6kZcAAAAAElFTkSuQmCC");
        using var ms = new MemoryStream(pngBytes);
        Shape picture = builder.InsertImage(ms);
        picture.WrapType = WrapType.None;
        picture.Width = 100;
        picture.Height = 100;

        // ---------- Insert a textbox (implemented as a rectangle with text) ----------
        Shape textBox = builder.InsertShape(ShapeType.Rectangle, 120, 40);
        textBox.WrapType = WrapType.None;
        textBox.Fill.Color = System.Drawing.Color.Transparent;
        textBox.Stroke.Color = System.Drawing.Color.Black;
        textBox.AppendChild(new Paragraph(doc));
        textBox.FirstParagraph.AppendChild(new Run(doc, "Hello World!"));

        // ---------- Insert an AutoShape (rectangle) ----------
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 80, 60);
        rectangle.Stroke.Color = System.Drawing.Color.Blue;

        // ---------- Group the shapes ----------
        GroupShape group = builder.InsertGroupShape(picture, textBox, rectangle);

        // Reposition child shapes within the group.
        picture.Left = 0;
        picture.Top = 0;

        textBox.Left = picture.Width + 10;
        textBox.Top = 0;

        rectangle.Left = 0;
        rectangle.Top = picture.Height + 10;

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "GroupShapeExample.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
