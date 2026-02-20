using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class AddShapesToGroup
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an empty paragraph to host the group shape.
        builder.Writeln();

        // Create a GroupShape and set its size and coordinate space.
        GroupShape group = new GroupShape(doc);
        group.Width = 400;               // Width in points.
        group.Height = 200;              // Height in points.
        group.CoordOrigin = new System.Drawing.Point(0, 0);
        group.CoordSize = new System.Drawing.Size(400, 200);

        // Add the group shape to the current paragraph.
        builder.CurrentParagraph.AppendChild(group);

        // ----- Add child shapes to the group -----

        // 1. Rectangle shape.
        Shape rect = new Shape(doc, ShapeType.Rectangle);
        rect.Width = 100;
        rect.Height = 80;
        rect.Left = 20;   // Position relative to the group's coordinate space.
        rect.Top = 20;
        rect.Fill.Color = System.Drawing.Color.LightBlue;
        group.AppendChild(rect);

        // 2. Ellipse shape.
        Shape ellipse = new Shape(doc, ShapeType.Ellipse);
        ellipse.Width = 80;
        ellipse.Height = 80;
        ellipse.Left = 150;
        ellipse.Top = 30;
        ellipse.Fill.Color = System.Drawing.Color.LightCoral;
        group.AppendChild(ellipse);

        // 3. TextBox shape.
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.Width = 120;
        textBox.Height = 60;
        textBox.Left = 260;
        textBox.Top = 100;
        // Add a paragraph with a run to provide the text inside the textbox.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Hello Group!");
        para.AppendChild(run);
        textBox.AppendChild(para);
        // Set internal margins for the textbox.
        textBox.TextBox.InternalMarginTop = 5;
        textBox.TextBox.InternalMarginBottom = 5;
        textBox.TextBox.InternalMarginLeft = 5;
        textBox.TextBox.InternalMarginRight = 5;
        group.AppendChild(textBox);

        // Save the document to a DOCX file.
        doc.Save("GroupShapeWithChildren.docx");
    }
}
