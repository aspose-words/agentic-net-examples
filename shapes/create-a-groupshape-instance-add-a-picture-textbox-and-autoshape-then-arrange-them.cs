using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class GroupShapeExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a group shape that will contain other shapes.
        GroupShape group = new GroupShape(doc);
        // Size and position of the group on the page (points).
        group.Bounds = new System.Drawing.RectangleF(0, 0, 400, 300);
        // Internal coordinate system of the group.
        group.CoordSize = new System.Drawing.Size(500, 500);
        group.CoordOrigin = new System.Drawing.Point(0, 0);

        // ---------- Picture shape ----------
        Shape picture = new Shape(doc, ShapeType.Image)
        {
            Width = 100,
            Height = 100,
            Left = 20,
            Top = 20,
            WrapType = WrapType.None
        };

        // Tiny 1x1 transparent PNG encoded in Base64.
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // Use a stream overload – this works with all Aspose.Words versions.
        using (MemoryStream ms = new MemoryStream(imageBytes))
        {
            picture.ImageData.SetImage(ms);
        }

        group.AppendChild(picture);

        // ---------- Text box ----------
        Shape textBox = new Shape(doc, ShapeType.TextBox)
        {
            Width = 150,
            Height = 50,
            Left = 150,
            Top = 20,
            WrapType = WrapType.None
        };
        // Add a paragraph with some text inside the text box.
        textBox.AppendChild(new Paragraph(doc));
        Paragraph tbParagraph = textBox.FirstParagraph;
        tbParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        Run tbRun = new Run(doc) { Text = "Hello Group!" };
        tbParagraph.AppendChild(tbRun);
        group.AppendChild(textBox);

        // ---------- AutoShape (rectangle) ----------
        Shape rectangle = new Shape(doc, ShapeType.Rectangle)
        {
            Width = 120,
            Height = 80,
            Left = 20,
            Top = 150,
            FillColor = System.Drawing.Color.Yellow,
            Stroke = { Color = System.Drawing.Color.Red }
        };
        group.AppendChild(rectangle);

        // Insert the completed group shape into the document.
        builder.InsertNode(group);

        // Save the document.
        string outputPath = "GroupShapeExample.docx";
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
