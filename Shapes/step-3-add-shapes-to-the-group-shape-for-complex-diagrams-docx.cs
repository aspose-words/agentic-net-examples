using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class AddShapesToGroupShape
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a group shape that will contain other shapes.
        GroupShape group = new GroupShape(doc);
        // Set the size of the group shape (in points).
        group.Width = 400;
        group.Height = 300;
        // Position the group shape on the page.
        group.Left = 100;
        group.Top = 100;
        // Define the coordinate space inside the group shape.
        group.CoordOrigin = new Point(0, 0);
        group.CoordSize = new Size(5000, 5000); // Large enough for child shapes.

        // -------------------------------------------------
        // Add a rectangle shape to the group.
        Shape rect = new Shape(doc, ShapeType.Rectangle);
        rect.Width = 200;
        rect.Height = 100;
        rect.Left = 50;   // Position relative to the group's coordinate space.
        rect.Top = 50;
        rect.Fill.ForeColor = Color.LightBlue;
        rect.Fill.Visible = true;
        rect.Stroke.Color = Color.DarkBlue;
        rect.Stroke.Weight = 2.0;
        group.AppendChild(rect);

        // Add a line shape to the group.
        Shape line = new Shape(doc, ShapeType.Line);
        line.Width = 300;
        line.Height = 0; // Height is not used for a horizontal line.
        line.Left = 100;
        line.Top = 200;
        line.Stroke.Color = Color.Red;
        line.Stroke.Weight = 3.0;
        line.Stroke.DashStyle = DashStyle.Dash;
        line.Stroke.StartArrowType = ArrowType.Arrow;
        line.Stroke.EndArrowType = ArrowType.Diamond;
        group.AppendChild(line);

        // Add a text box shape to the group.
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.Width = 180;
        textBox.Height = 60;
        textBox.Left = 250;
        textBox.Top = 150;
        textBox.Fill.ForeColor = Color.LightYellow;
        textBox.Fill.Visible = true;
        textBox.Stroke.Color = Color.Orange;
        textBox.Stroke.Weight = 1.5;
        // Insert some text into the text box.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Group Shape Text");
        para.AppendChild(run);
        textBox.FirstParagraph.AppendChild(para);
        group.AppendChild(textBox);
        // -------------------------------------------------

        // Insert the group shape into the document.
        builder.InsertNode(group);

        // Save the document to a DOCX file.
        doc.Save("GroupShapeComplex.docx");
    }
}
