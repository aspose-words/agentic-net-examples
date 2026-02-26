using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating text box shape with a width of 200 points and a height of 100 points.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);
        // Prevent the text box from affecting surrounding text layout.
        textBox.WrapType = WrapType.None;

        // Move the builder's cursor inside the text box so we can add content.
        builder.MoveTo(textBox.LastParagraph);
        builder.Writeln("Hello world!");
        builder.Write("This is a text box.");

        // Save the document to a DOCX file.
        doc.Save("TextBoxExample.docx");
    }
}
