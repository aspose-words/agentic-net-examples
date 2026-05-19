using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeWordsTextBoxExample
{
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

            // Insert the text box into the document.
            builder.InsertNode(textBox);

            // Add an empty paragraph to the text box – this will hold the text.
            textBox.AppendChild(new Paragraph(doc));

            // Move the builder cursor to the first paragraph inside the text box.
            builder.MoveTo(textBox.FirstParagraph);

            // Set the paragraph direction to right‑to‑left.
            builder.ParagraphFormat.Bidi = true;

            // Ensure the font treats the text as right‑to‑left.
            builder.Font.Bidi = true;

            // Insert Arabic text.
            builder.Write("مرحبا بالعالم"); // "Hello World" in Arabic.

            // Prepare output folder.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Save the document.
            string outputPath = Path.Combine(outputDir, "TextBox_RTL_Arabic.docx");
            doc.Save(outputPath);
        }
    }
}
