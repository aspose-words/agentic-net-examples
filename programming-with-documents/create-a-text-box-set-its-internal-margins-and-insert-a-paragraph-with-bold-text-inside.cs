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

        // Insert a textbox shape (width: 200 points, height: 100 points).
        Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 200, 100);

        // Set internal margins of the textbox (in points).
        TextBox textBox = textBoxShape.TextBox;
        textBox.InternalMarginTop = 10;
        textBox.InternalMarginBottom = 10;
        textBox.InternalMarginLeft = 10;
        textBox.InternalMarginRight = 10;

        // Move the builder cursor inside the textbox.
        builder.MoveTo(textBoxShape.LastParagraph);

        // Insert a paragraph with bold text.
        builder.Font.Bold = true;
        builder.Writeln("Bold text inside the textbox.");

        // Save the document.
        doc.Save("TextBoxMargins.docx");
    }
}
