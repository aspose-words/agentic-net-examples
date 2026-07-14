using System;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder attached to the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a textbox shape with a specific size.
            Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 200, 100);

            // Set the internal margins of the textbox (values are in points).
            TextBox textBox = textBoxShape.TextBox;
            textBox.InternalMarginTop = 15;
            textBox.InternalMarginBottom = 15;
            textBox.InternalMarginLeft = 15;
            textBox.InternalMarginRight = 15;

            // Move the builder's cursor to the inside of the textbox.
            builder.MoveTo(textBoxShape.LastParagraph);

            // Apply bold formatting to the text that will be inserted.
            builder.Font.Bold = true;

            // Insert a paragraph with bold text inside the textbox.
            builder.Writeln("This is bold text inside a textbox.");

            // Save the document to a file.
            doc.Save("TextBoxMargins.docx");
        }
    }
}
