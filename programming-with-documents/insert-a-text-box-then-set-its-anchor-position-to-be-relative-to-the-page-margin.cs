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

        // Insert a text box shape with a specific size.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);

        // Add some sample text inside the text box.
        // The TextBoxContentBuilder property is not available in this version,
        // so we add a paragraph and a run directly to the shape's text box.
        Paragraph paragraph = new Paragraph(doc);
        Run run = new Run(doc, "Hello, Aspose.Words!");
        paragraph.AppendChild(run);
        textBox.AppendChild(paragraph);

        // Set the anchor position of the text box relative to the page margin.
        textBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;
        textBox.RelativeVerticalPosition = RelativeVerticalPosition.Margin;

        // Position the text box 50 points from the left and top margins.
        textBox.Left = 50;
        textBox.Top = 50;

        // Save the document to a file.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
