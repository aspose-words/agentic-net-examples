using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class GroupShapeExample
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Create a picture shape ----------
        // Use a 1x1 pixel PNG (light blue) encoded in Base64.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X6ZcAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // Insert the image from a memory stream.
        using (MemoryStream ms = new MemoryStream(imageBytes))
        {
            Shape pictureShape = builder.InsertImage(ms);
            pictureShape.Name = "PictureShape";
            pictureShape.Width = 100;   // Desired display size.
            pictureShape.Height = 100;
            pictureShape.Left = 0;
            pictureShape.Top = 0;
            pictureShape.WrapType = WrapType.None; // Floating for positioning.
        }

        // The picture shape is the last inserted node.
        Shape picture = (Shape)doc.GetChild(NodeType.Shape, 0, true);

        // ---------- Create a textbox shape ----------
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 150, 50);
        textBox.Name = "TextBoxShape";
        textBox.WrapType = WrapType.None;
        textBox.Left = 120;
        textBox.Top = 0;

        // Add text to the textbox.
        Paragraph tbParagraph = textBox.FirstParagraph;
        tbParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        Run tbRun = new Run(doc, "Hello Group!");
        tbParagraph.AppendChild(tbRun);

        // ---------- Create an AutoShape (rectangle) ----------
        Shape autoShape = builder.InsertShape(ShapeType.Rectangle, 100, 100);
        autoShape.Name = "AutoShape";
        autoShape.WrapType = WrapType.None;
        autoShape.Left = 0;
        autoShape.Top = 120;
        autoShape.FillColor = System.Drawing.Color.LightCoral; // Using System.Drawing.Color is allowed for color values.

        // ---------- Group the three shapes ----------
        // InsertGroupShape removes the original shapes and inserts a new GroupShape.
        GroupShape group = builder.InsertGroupShape(picture, textBox, autoShape);
        group.Name = "MyGroupShape";

        // Optionally adjust the group's position.
        group.Left = 50;
        group.Top = 50;
        group.WrapType = WrapType.None;

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "GroupShapeExample.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
