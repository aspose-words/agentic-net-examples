using System;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeWordsTextBoxRtlExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a text box shape.
            Shape textBox = new Shape(doc, ShapeType.TextBox)
            {
                Width = 300,
                Height = 100
            };

            // Add an empty paragraph to the text box so we can write into it.
            textBox.AppendChild(new Paragraph(doc));

            // Insert the text box into the document.
            builder.InsertNode(textBox);

            // Move the builder's cursor to the first paragraph inside the text box.
            builder.MoveTo(textBox.FirstParagraph);

            // Set the text direction to right‑to‑left.
            builder.Font.Bidi = true;

            // Write Arabic text.
            builder.Write("مرحبا بالعالم!"); // "Hello world!" in Arabic.

            // Save the document to a file in the same folder as the executable.
            doc.Save("TextBoxRTL.docx");
        }
    }
}
