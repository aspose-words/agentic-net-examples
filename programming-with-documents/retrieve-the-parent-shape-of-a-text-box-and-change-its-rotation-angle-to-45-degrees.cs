using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text box shape.
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.Width = 200;
        textBox.Height = 100;
        textBox.WrapType = WrapType.Inline;

        // Add the shape to the document.
        builder.InsertNode(textBox);

        // Ensure the text box contains a paragraph to host text.
        Paragraph innerParagraph = new Paragraph(doc);
        textBox.AppendChild(innerParagraph);

        // Move the builder to the paragraph inside the text box and add text.
        builder.MoveTo(innerParagraph);
        builder.Writeln("Sample text inside the text box.");

        // Retrieve the paragraph that is inside the text box (first paragraph).
        Paragraph retrievedParagraph = (Paragraph)textBox.GetChildNodes(NodeType.Paragraph, true)[0];

        // The parent of this paragraph is the Shape (the text box itself).
        Shape parentShape = retrievedParagraph.ParentNode as Shape;

        // Change the rotation angle of the parent shape to 45 degrees.
        if (parentShape != null)
        {
            parentShape.Rotation = 45;
        }

        // Save the document.
        doc.Save("RotatedTextBox.docx");
    }
}
