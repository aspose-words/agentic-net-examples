using System;
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

        // Insert a regular rectangle shape (just to have another shape in the document).
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        rectangle.FillColor = Color.Yellow;

        // Insert a text box shape with some initial fill color.
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.WrapType = WrapType.None;
        textBox.Height = 50;
        textBox.Width = 200;
        textBox.FillColor = Color.LightBlue; // initial color

        // Add a paragraph with text inside the text box.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Sample text box");
        para.AppendChild(run);
        textBox.AppendChild(para);

        // Place the text box into the document body.
        builder.CurrentParagraph.AppendChild(textBox);

        // Iterate through all shapes in the document.
        NodeCollection allShapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in allShapes)
        {
            // If the shape is a text box, change its fill color to light gray.
            if (shape.ShapeType == ShapeType.TextBox)
            {
                shape.FillColor = Color.LightGray;
                // Alternatively: shape.Fill.Solid(Color.LightGray);
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
