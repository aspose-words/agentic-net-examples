using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using TextBox = Aspose.Words.Drawing.TextBox; // Alias to avoid conflict with System.Windows.Forms

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a group shape that will contain other shapes.
        GroupShape group = new GroupShape(doc);
        // Set the size and position of the group shape.
        group.Width = 300;   // Width in points.
        group.Height = 150;  // Height in points.
        group.Left = 100;    // Distance from the left edge of the page.
        group.Top = 100;     // Distance from the top edge of the page.
        group.WrapType = WrapType.None; // No text wrapping.
        group.BehindText = true;        // Place behind the document text.

        // -------------------------------------------------
        // Add a rectangle shape inside the group.
        // -------------------------------------------------
        Shape rectangle = new Shape(doc, ShapeType.Rectangle);
        rectangle.Width = 120;
        rectangle.Height = 80;
        rectangle.Left = 0;   // Position relative to the group's coordinate space.
        rectangle.Top = 0;
        rectangle.Fill.Color = System.Drawing.Color.LightBlue;
        rectangle.Stroked = true;
        rectangle.StrokeColor = System.Drawing.Color.DarkBlue;
        rectangle.StrokeWeight = 1.0;
        group.AppendChild(rectangle);

        // -------------------------------------------------
        // Add a text box shape inside the group.
        // -------------------------------------------------
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.Width = 150;
        textBox.Height = 60;
        textBox.Left = 130; // Position next to the rectangle.
        textBox.Top = 0;
        // Configure internal margins for the text box.
        textBox.TextBox.InternalMarginTop = 5;
        textBox.TextBox.InternalMarginBottom = 5;
        textBox.TextBox.InternalMarginLeft = 5;
        textBox.TextBox.InternalMarginRight = 5;
        // Add a paragraph with some text to the text box.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Grouped Text");
        para.AppendChild(run);
        textBox.AppendChild(para);
        group.AppendChild(textBox);

        // Insert the group shape into the document at the current builder position.
        builder.InsertNode(group);

        // Save the document to a file.
        doc.Save("GroupShape.docx");
    }
}
