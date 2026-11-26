using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text box shape with some initial fill color.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);
        textBox.FillColor = Color.Yellow; // Initial fill color for demonstration.
        textBox.AppendChild(new Paragraph(doc));
        textBox.FirstParagraph.AppendChild(new Run(doc, "Sample text inside the text box."));

        // Iterate through all shapes in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            // Check if the shape is a text box.
            if (shape.ShapeType == ShapeType.TextBox)
            {
                // Change the fill color to light gray.
                shape.FillColor = Color.LightGray;
            }
        }

        // Save the modified document.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
