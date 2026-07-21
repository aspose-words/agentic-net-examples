using System;
using System.Drawing;
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

        // -----------------------------------------------------------------
        // 1. Create a picture shape from an in‑memory PNG image.
        // -----------------------------------------------------------------
        // This is a 1×1 pixel PNG (any image will work for the example).
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // Create an image shape (do not insert it yet – it will be added to the group later).
        Shape pictureShape = new Shape(doc, ShapeType.Image);
        using (MemoryStream ms = new MemoryStream(imageBytes))
        {
            pictureShape.ImageData.SetImage(ms);
        }
        pictureShape.Width = 100;   // 100 points
        pictureShape.Height = 100;  // 100 points
        pictureShape.WrapType = WrapType.None;
        pictureShape.Left = 0;
        pictureShape.Top = 0;

        // -----------------------------------------------------------------
        // 2. Create a textbox shape.
        // -----------------------------------------------------------------
        Shape textBox = new Shape(doc, ShapeType.TextBox)
        {
            Width = 150,
            Height = 50,
            Left = 120,
            Top = 0,
            WrapType = WrapType.None
        };
        // Add a paragraph with some text inside the textbox.
        Paragraph tbParagraph = new Paragraph(doc);
        Run tbRun = new Run(doc, "Hello Aspose!");
        tbParagraph.AppendChild(tbRun);
        textBox.AppendChild(tbParagraph);

        // -----------------------------------------------------------------
        // 3. Create an AutoShape (a simple rectangle).
        // -----------------------------------------------------------------
        Shape autoShape = new Shape(doc, ShapeType.Rectangle)
        {
            Width = 80,
            Height = 80,
            Left = 0,
            Top = 120,
            WrapType = WrapType.None
        };

        // -----------------------------------------------------------------
        // 4. Create a GroupShape and add the three shapes to it.
        // -----------------------------------------------------------------
        GroupShape group = new GroupShape(doc);
        // Define the group's outer bounds (position and size in the document).
        group.Bounds = new RectangleF(50, 50, 300, 300);
        // Define the internal coordinate system (optional but makes positioning easier).
        group.CoordSize = new Size(500, 500);
        group.CoordOrigin = new Point(0, 0);

        // Append child shapes to the group.
        group.AppendChild(pictureShape);
        group.AppendChild(textBox);
        group.AppendChild(autoShape);

        // Insert the group shape into the document at the current builder position.
        builder.InsertNode(group);

        // Save the document.
        doc.Save("GroupShapeExample.docx");
    }
}
