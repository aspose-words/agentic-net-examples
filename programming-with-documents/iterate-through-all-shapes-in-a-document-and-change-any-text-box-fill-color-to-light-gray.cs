using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text box shape with an initial fill color (e.g., White).
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);
        textBox.FillColor = Color.White;
        textBox.AppendChild(new Paragraph(doc));
        textBox.FirstParagraph.AppendChild(new Run(doc, "Sample text"));

        // Iterate through all shapes and change fill color of text boxes to LightGray.
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

        // Verify that the file was saved.
        if (File.Exists(outputPath))
        {
            // Optionally, reload to ensure no exceptions.
            Document loadedDoc = new Document(outputPath);
        }
    }
}
