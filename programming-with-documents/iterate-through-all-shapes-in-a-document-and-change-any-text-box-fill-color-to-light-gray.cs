using System;
using System.IO;
using System.Linq;
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

        // Add some regular text.
        builder.Writeln("Document with a text box shape.");

        // Create a text box shape with an initial fill color.
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.Width = 200;
        textBox.Height = 50;
        textBox.WrapType = WrapType.None;
        textBox.FillColor = Color.Yellow; // initial color

        // Add a paragraph and a run inside the text box.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Text inside the box");
        para.AppendChild(run);
        textBox.AppendChild(para);

        // Insert the text box into the document.
        builder.InsertNode(textBox);

        // Iterate through all shapes in the document.
        NodeCollection allShapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in allShapes.OfType<Shape>())
        {
            // Change fill color only for text box shapes.
            if (shape.ShapeType == ShapeType.TextBox)
            {
                shape.FillColor = Color.LightGray;
                // Alternatively: shape.Fill.Solid(Color.LightGray);
            }
        }

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);
    }
}
