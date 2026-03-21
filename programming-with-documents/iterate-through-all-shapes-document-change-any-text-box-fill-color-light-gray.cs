using System;
using System.Drawing;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a textbox shape to the document so we have something to modify.
        Shape textBox = new Shape(doc, ShapeType.TextBox)
        {
            Width = 200,
            Height = 100,
            WrapType = WrapType.Inline,
            Fill = { Color = Color.White } // initial fill color
        };

        Paragraph para = new Paragraph(doc);
        para.AppendChild(textBox);
        doc.FirstSection.Body.AppendChild(para);

        // Get all shapes in the document (including those inside text boxes).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate through each shape.
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            // Check if the shape is a text box.
            if (shape.ShapeType == ShapeType.TextBox)
            {
                // Change the fill to a solid light gray color.
                shape.Fill.Solid(Color.LightGray);
            }
        }

        // Save the modified document.
        doc.Save("output.docx");
    }
}
