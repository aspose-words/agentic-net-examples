using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a text box shape.
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.Width = 200;   // Width in points.
        textBox.Height = 100;  // Height in points.

        // Set internal margins of the text box (in points).
        textBox.TextBox.InternalMarginTop = 5;
        textBox.TextBox.InternalMarginBottom = 5;
        textBox.TextBox.InternalMarginLeft = 5;
        textBox.TextBox.InternalMarginRight = 5;

        // Insert the text box into the document.
        builder.InsertNode(textBox);

        // Create a paragraph with bold text.
        Paragraph paragraph = new Paragraph(doc);
        Run run = new Run(doc, "Bold text inside the text box");
        run.Font.Bold = true;
        paragraph.AppendChild(run);

        // Add the paragraph to the text box.
        textBox.AppendChild(paragraph);

        // Save the document to a file.
        doc.Save("Output.docx");
    }
}
