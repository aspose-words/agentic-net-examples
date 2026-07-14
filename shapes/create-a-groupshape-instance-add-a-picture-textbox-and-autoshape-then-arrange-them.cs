using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class GroupShapeExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Create a sample image (1x1 pixel PNG) ----------
        // The image data is embedded as a Base64 string to avoid external dependencies.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");

        // ---------- Create a picture shape ----------
        Shape pictureShape = new Shape(doc, ShapeType.Image);
        pictureShape.ImageData.SetImage(new MemoryStream(pngBytes));
        pictureShape.Width = 100;
        pictureShape.Height = 100;
        pictureShape.Left = 0;
        pictureShape.Top = 0;
        pictureShape.WrapType = WrapType.None;

        // ---------- Create a text box shape ----------
        Shape textBoxShape = new Shape(doc, ShapeType.TextBox);
        textBoxShape.Width = 120;
        textBoxShape.Height = 60;
        textBoxShape.Left = 120; // Position to the right of the picture.
        textBoxShape.Top = 0;
        textBoxShape.WrapType = WrapType.None;
        textBoxShape.FillColor = Color.LightYellow;
        textBoxShape.Stroke.Color = Color.Orange;

        // Add a paragraph with some text inside the text box.
        textBoxShape.AppendChild(new Paragraph(doc));
        Paragraph tbParagraph = textBoxShape.FirstParagraph;
        tbParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        Run tbRun = new Run(doc) { Text = "Hello World!" };
        tbParagraph.AppendChild(tbRun);

        // ---------- Create an AutoShape (rectangle) ----------
        Shape autoShape = new Shape(doc, ShapeType.Rectangle);
        autoShape.Width = 150;
        autoShape.Height = 80;
        autoShape.Left = 0;
        autoShape.Top = 120; // Position below the picture.
        autoShape.WrapType = WrapType.None;
        autoShape.FillColor = Color.LightGreen;
        autoShape.Stroke.Color = Color.DarkGreen;

        // ---------- Create a GroupShape and add the three shapes ----------
        GroupShape group = new GroupShape(doc);
        // Define the outer bounds of the group (in points).
        group.Bounds = new RectangleF(0, 0, 300, 250);
        // Define the internal coordinate system of the group.
        group.CoordSize = new Size(500, 500);
        group.CoordOrigin = new Point(0, 0);

        // Append child shapes to the group.
        group.AppendChild(pictureShape);
        group.AppendChild(textBoxShape);
        group.AppendChild(autoShape);

        // Insert the group into the document at the current builder position.
        builder.InsertNode(group);

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "GroupShapeExample.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }
}
