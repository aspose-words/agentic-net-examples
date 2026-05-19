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

        // -----------------------------------------------------------------
        // 1. Create a simple 1x1 PNG image in memory (transparent pixel).
        // -----------------------------------------------------------------
        // Base64 representation of a 1x1 transparent PNG.
        byte[] pngData = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=");

        // -----------------------------------------------------------------
        // 2. Create the individual shapes that will be placed inside the group.
        // -----------------------------------------------------------------

        // Picture shape (inserted via InsertImage which accepts a byte array).
        Shape picture = builder.InsertImage(pngData);
        picture.Width = 100;
        picture.Height = 100;
        picture.Left = 0;
        picture.Top = 0;
        picture.WrapType = WrapType.None; // Floating picture.

        // Text box shape.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 120, 50);
        textBox.Left = 110; // Position to the right of the picture.
        textBox.Top = 0;
        textBox.WrapType = WrapType.None;

        // Add a paragraph with some text inside the text box.
        Paragraph tbParagraph = new Paragraph(doc);
        Run tbRun = new Run(doc, "Hello Aspose!");
        tbParagraph.AppendChild(tbRun);
        textBox.AppendChild(tbParagraph);

        // AutoShape (a simple rectangle).
        Shape autoShape = builder.InsertShape(ShapeType.Rectangle, 100, 80);
        autoShape.Left = 0;
        autoShape.Top = 110; // Position below the picture.
        autoShape.FillColor = Color.LightCoral;
        autoShape.Stroke.Color = Color.DarkRed;

        // -----------------------------------------------------------------
        // 3. Group the three shapes.
        // -----------------------------------------------------------------
        // InsertGroupShape automatically calculates the group bounds and inserts the group.
        GroupShape group = builder.InsertGroupShape(picture, textBox, autoShape);

        // -----------------------------------------------------------------
        // 4. Save the document.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "GroupShapeExample.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }
}
