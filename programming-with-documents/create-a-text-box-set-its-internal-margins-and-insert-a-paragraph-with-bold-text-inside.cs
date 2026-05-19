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

        // Insert a textbox shape.
        Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 200, 100);
        TextBox textBox = textBoxShape.TextBox;

        // Set internal margins (in points).
        textBox.InternalMarginTop = 15;
        textBox.InternalMarginBottom = 15;
        textBox.InternalMarginLeft = 15;
        textBox.InternalMarginRight = 15;

        // Move the builder cursor inside the textbox.
        builder.MoveTo(textBoxShape.LastParagraph);

        // Set bold formatting and add a paragraph with bold text.
        builder.Font.Bold = true;
        builder.Writeln("Bold text inside the textbox");

        // Save the document.
        doc.Save("TextBoxMargins.docx");
    }
}
