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
        Shape textBox = new Shape(doc, ShapeType.TextBox)
        {
            Width = 300,
            Height = 100
        };

        // Add an empty paragraph to the textbox – this will hold the text.
        Paragraph paragraph = new Paragraph(doc);
        textBox.AppendChild(paragraph);

        // Insert the textbox into the document.
        builder.InsertNode(textBox);

        // Move the builder's cursor to the paragraph inside the textbox.
        builder.MoveTo(paragraph);

        // Set the paragraph direction to right‑to‑left.
        paragraph.ParagraphFormat.Bidi = true;

        // Write Arabic text.
        builder.Write("مرحبا بالعالم");

        // Save the document.
        doc.Save("TextBoxRTL.docx");
    }
}
