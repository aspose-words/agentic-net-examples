using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class AddShapesToGroupShape
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a group shape that will hold child shapes.
        GroupShape group = new GroupShape(doc);

        // Set the size and position of the group shape.
        group.Bounds = new RectangleF(50, 50, 400, 300); // left, top, width, height

        // Optional: configure the internal coordinate system of the group.
        group.CoordSize = new Size(1000, 1000); // default, but shown for clarity
        group.CoordOrigin = new Point(0, 0);

        // ----- Create child shapes -----

        // Rectangle shape.
        Shape rect = new Shape(doc, ShapeType.Rectangle)
        {
            Width = 200,
            Height = 150,
            Left = 100,   // position within the group's coordinate space
            Top = 75,
            FillColor = Color.LightBlue,
            Stroke = { Color = Color.DarkBlue, Weight = 1.5 }
        };

        // Ellipse shape.
        Shape ellipse = new Shape(doc, ShapeType.Ellipse)
        {
            Width = 150,
            Height = 150,
            Left = 250,
            Top = 100,
            FillColor = Color.LightCoral,
            Stroke = { Color = Color.Maroon, Weight = 1.5 }
        };

        // Text box shape.
        Shape textBox = new Shape(doc, ShapeType.TextBox)
        {
            Width = 180,
            Height = 60,
            Left = 150,
            Top = 200,
            FillColor = Color.White,
            Stroke = { Color = Color.Gray, Weight = 1 }
        };

        // Add a paragraph inside the text box.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Hello GroupShape!");
        para.AppendChild(run);
        textBox.AppendChild(para);

        // ----- Append child shapes to the group -----
        group.AppendChild(rect);
        group.AppendChild(ellipse);
        group.AppendChild(textBox);

        // Insert the group shape into the document at the current builder position.
        builder.InsertNode(group);

        // Save the document.
        doc.Save("AddChildShapesToGroupShape.docx");
    }
}
