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

        // Create a textbox shape.
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.Width = 300;
        textBox.Height = 100;

        // Create a paragraph with Arabic text.
        Paragraph paragraph = new Paragraph(doc);
        Run run = new Run(doc, "مرحبا بالعالم");
        paragraph.AppendChild(run);

        // Set the paragraph direction to right‑to‑left.
        paragraph.ParagraphFormat.Bidi = true;

        // Add the paragraph to the textbox.
        textBox.AppendChild(paragraph);

        // Insert the textbox into the document.
        builder.InsertNode(textBox);

        // Save the document.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
