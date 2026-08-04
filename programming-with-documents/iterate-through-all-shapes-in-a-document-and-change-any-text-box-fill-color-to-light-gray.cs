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

        // Insert some regular text.
        builder.Writeln("Document before the text box.");

        // Create a floating text box shape.
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.WrapType = WrapType.None;
        textBox.Height = 100;
        textBox.Width = 200;

        // Add a paragraph with a run of text inside the text box.
        textBox.AppendChild(new Paragraph(doc));
        Paragraph tbParagraph = (Paragraph)textBox.FirstChild;
        Run tbRun = new Run(doc, "This is a text box.");
        tbParagraph.AppendChild(tbRun);

        // Add the text box to the document.
        builder.CurrentParagraph.AppendChild(textBox);
        builder.Writeln("Document after the text box.");

        // Insert another shape (rectangle) to demonstrate that non‑textbox shapes are unchanged.
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 120, 60);
        rectangle.FillColor = Color.Blue;

        // Iterate through all shapes in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            // Change fill color only for text box shapes.
            if (shape.ShapeType == ShapeType.TextBox)
            {
                shape.FillColor = Color.LightGray;
            }
        }

        // Save the modified document.
        const string outputPath = "Result.docx";
        doc.Save(outputPath);
    }
}
