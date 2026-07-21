using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a textbox shape.
        Shape textBox = new Shape(doc, ShapeType.TextBox)
        {
            Width = 300,
            Height = 100
        };

        // Add an empty paragraph to the textbox – this is where the text will go.
        Paragraph para = new Paragraph(doc);
        textBox.AppendChild(para);

        // Insert the textbox into the document.
        builder.InsertNode(textBox);

        // Move the builder's cursor to the first paragraph inside the textbox.
        builder.MoveTo(textBox.FirstParagraph);

        // Set the paragraph to be right‑to‑left.
        builder.ParagraphFormat.Bidi = true;

        // Set the font to treat the run as right‑to‑left.
        builder.Font.Bidi = true;

        // Arabic text to insert.
        string arabic = "مرحبا بالعالم!";

        // Write the Arabic text into the textbox.
        builder.Write(arabic);

        // Save the document.
        doc.Save("TextBox_RTL_Arabic.docx");
    }
}
