using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a textbox shape with specified width and height.
        Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 300, 150);

        // Access the TextBox object to configure its properties.
        TextBox textBox = textBoxShape.TextBox;
        textBox.TextBoxWrapMode = TextBoxWrapMode.Square; // Wrap text inside the box.
        textBox.VerticalAnchor = TextBoxAnchor.Middle;   // Vertically center the text.

        // Move the builder cursor inside the textbox and add some text.
        builder.MoveTo(textBoxShape.FirstParagraph);
        builder.Font.Size = 24;
        builder.Font.Bold = true;
        builder.Write("Hello Aspose.Words TextBox!");

        // Save the document to a DOCX file.
        doc.Save("TextBoxExample.docx");
    }
}
