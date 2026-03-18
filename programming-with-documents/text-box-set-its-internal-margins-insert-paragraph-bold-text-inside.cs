using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document and attach a DocumentBuilder to it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating text box shape with a width of 200 points and a height of 100 points.
        Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 200, 100);

        // Access the TextBox object of the shape to configure its internal margins (in points).
        TextBox textBox = textBoxShape.TextBox;
        textBox.InternalMarginTop = 15;
        textBox.InternalMarginBottom = 15;
        textBox.InternalMarginLeft = 15;
        textBox.InternalMarginRight = 15;

        // Move the builder's cursor to the last paragraph inside the text box.
        builder.MoveTo(textBoxShape.LastParagraph);

        // Apply bold formatting to the text that will be inserted.
        builder.Font.Bold = true;

        // Insert a paragraph containing bold text.
        builder.Writeln("This is bold text inside the text box.");

        // Save the resulting document to a file.
        doc.Save("TextBoxWithBoldParagraph.docx");
    }
}
