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

        // Insert a text box shape.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);

        // Add sample text inside the text box.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Sample text inside the text box.");
        para.AppendChild(run);
        textBox.AppendChild(para);

        // Iterate through all shapes and set fill color of text boxes to light gray.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            if (shape.ShapeType == ShapeType.TextBox)
            {
                shape.FillColor = Color.LightGray;
            }
        }

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
