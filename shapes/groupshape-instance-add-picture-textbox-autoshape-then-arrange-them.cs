using System;
using System.Drawing;
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
        // Insert a floating image (replace with a valid image path on your machine).
        Shape picture = builder.InsertImage(@"C:\Images\Sample.jpg");
        picture.WrapType = WrapType.None;               // No text wrapping.
        picture.Width = 100;                             // Set desired size.
        picture.Height = 100;

        // ---------- Insert a textbox ----------
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.Width = 120;
        textBox.Height = 40;
        textBox.WrapType = WrapType.None;
        // Add a paragraph with some text inside the textbox.
        textBox.AppendChild(new Paragraph(doc));
        Paragraph tbParagraph = textBox.FirstParagraph;
        Run tbRun = new Run(doc);
        tbRun.Text = "Hello World!";
        tbParagraph.AppendChild(tbRun);

        // ---------- Insert an AutoShape (rectangle) ----------
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 80, 60);
        rectangle.Stroke.Color = Color.Blue;            // Outline color.

        // ---------- Group the shapes ----------
        // The InsertGroupShape method automatically calculates the group's bounds.
        GroupShape group = builder.InsertGroupShape(picture, textBox, rectangle);

        // Optional: reposition child shapes within the group.
        // Coordinates are relative to the group's internal coordinate system.
        picture.Left = 0;
        picture.Top = 0;

        textBox.Left = picture.Width + 10;               // Place textbox to the right of the picture.
        textBox.Top = 0;

        rectangle.Left = 0;
        rectangle.Top = picture.Height + 10;             // Place rectangle below the picture.

        // Save the document.
        doc.Save(@"C:\Output\GroupShapeExample.docx");
    }
}
