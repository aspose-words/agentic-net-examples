using System;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace LinkedTextBoxExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert the first text box.
            Shape textBox1 = builder.InsertShape(ShapeType.TextBox, 300, 100);
            // Configure the text box to allow overflow (do not fit shape to text and no wrapping).
            textBox1.TextBox.FitShapeToText = false;
            textBox1.TextBox.TextBoxWrapMode = TextBoxWrapMode.None;

            // Move the builder inside the first text box and add a long paragraph.
            builder.MoveTo(textBox1.LastParagraph);
            string longText = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. ";
            // Repeat the text to ensure it exceeds the size of the first box.
            for (int i = 0; i < 30; i++)
                builder.Write(longText);

            // Insert a second text box that will receive the overflow text.
            // Move the cursor to the end of the document (after the first text box).
            builder.MoveToDocumentEnd();
            Shape textBox2 = builder.InsertShape(ShapeType.TextBox, 300, 100);
            textBox2.TextBox.FitShapeToText = false;
            textBox2.TextBox.TextBoxWrapMode = TextBoxWrapMode.None;

            // Link the first text box to the second one.
            // Overflow text from the first box will continue in the second box.
            textBox1.TextBox.Next = textBox2.TextBox;

            // Save the document.
            doc.Save("LinkedTextBox.docx");
        }
    }
}
