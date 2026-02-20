using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a group shape that will contain other shapes.
        GroupShape group = new GroupShape(doc);
        group.Width = 200;               // Width of the group.
        group.Height = 100;              // Height of the group.
        group.Left = 100;                // Position from the left of the page.
        group.Top = 100;                 // Position from the top of the page.
        group.WrapType = WrapType.None; // Make it a floating shape.
        group.BehindText = false;

        // Add a rectangle shape to the group.
        Shape rect = new Shape(doc, ShapeType.Rectangle);
        rect.Width = 200;
        rect.Height = 100;
        rect.Fill.ForeColor = Color.LightBlue;
        rect.Stroke.Color = Color.DarkBlue;
        rect.Left = 0;   // Position inside the group.
        rect.Top = 0;
        group.AppendChild(rect);

        // Add a line shape to the group.
        Shape line = new Shape(doc, ShapeType.Line);
        line.Width = 180;
        line.Height = 0;
        line.Left = 10;
        line.Top = 50;
        line.Stroke.Color = Color.Red;
        line.StrokeWeight = 2;
        group.AppendChild(line);

        // Insert the group shape into the document.
        builder.InsertNode(group);

        // Save the document as PDF, rendering DrawingML shapes (including the group) correctly.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            DmlRenderingMode = DmlRenderingMode.DrawingML
        };
        doc.Save("GroupShape.pdf", pdfOptions);
    }
}
