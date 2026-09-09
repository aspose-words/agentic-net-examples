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
        Shape textBox = new Shape(doc, ShapeType.TextBox)
        {
            Width = 300,
            Height = 100
        };

        // Set the layout flow (horizontal is the default).
        textBox.TextBox.LayoutFlow = LayoutFlow.Horizontal;

        // Add an empty paragraph to the text box – this is where the text will go.
        Paragraph paragraph = new Paragraph(doc);
        textBox.AppendChild(paragraph);

        // Insert the text box into the document.
        builder.InsertNode(textBox);

        // Move the builder's cursor to the paragraph inside the text box.
        builder.MoveTo(paragraph);

        // Set the paragraph direction to right‑to‑left.
        paragraph.ParagraphFormat.Bidi = true;

        // Write Arabic characters.
        builder.Write("مرحبا بالعالم!");

        // Save the document.
        doc.Save("TextBoxRTL.docx");
    }
}
