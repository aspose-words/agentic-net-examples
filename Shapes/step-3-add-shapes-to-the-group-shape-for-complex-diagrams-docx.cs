using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class AddShapesToGroupShape
{
    static void Main()
    {
        // Create a new blank document and a DocumentBuilder for editing.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two basic shapes that will be grouped together.
        Shape rect = builder.InsertShape(ShapeType.Rectangle, 200, 150);
        rect.Left = 50;
        rect.Top = 50;
        rect.Stroke.Color = Color.Blue;

        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 150);
        ellipse.Left = 300;
        ellipse.Top = 70;
        ellipse.Stroke.Color = Color.Green;

        // Group the rectangle and ellipse into a single GroupShape.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Add additional shapes directly to the existing group.
        // Example: a star shape positioned relative to the group's coordinate system.
        Shape star = new Shape(doc, ShapeType.Star)
        {
            Width = 80,
            Height = 80,
            Left = -40,   // Center the star within the group.
            Top = -40,
            FillColor = Color.Yellow,
            Stroke = { Color = Color.Orange }
        };
        group.AppendChild(star);

        // Example: a text box shape placed near the bottom‑right corner of the group.
        Shape textBox = new Shape(doc, ShapeType.TextBox)
        {
            Width = 120,
            Height = 40,
            // Position using the group's internal coordinate plane.
            Left = group.CoordSize.Width + group.CoordOrigin.X - 120,
            Top = group.CoordSize.Height + group.CoordOrigin.Y,
            FillColor = Color.LightGray
        };
        // Add a paragraph with text inside the text box.
        Paragraph para = new Paragraph(doc);
        para.AppendChild(new Run(doc, "Complex Diagram"));
        textBox.AppendChild(para);
        group.AppendChild(textBox);

        // Optionally adjust the group's internal coordinate system for better scaling.
        group.CoordSize = new Size(1000, 1000);
        group.CoordOrigin = new Point(0, 0);

        // Insert the completed group shape into the document at the current builder position.
        // (The group is already inserted by InsertGroupShape, so no further insertion is required.)

        // Save the document to a DOCX file.
        doc.Save("ComplexDiagramGroupShape.docx");
    }
}
