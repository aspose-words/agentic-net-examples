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

        // Insert a textbox shape with specific dimensions.
        Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 200, 100);

        // Set internal margins (in points) for the textbox.
        TextBox textBox = textBoxShape.TextBox;
        textBox.InternalMarginTop = 10;
        textBox.InternalMarginBottom = 10;
        textBox.InternalMarginLeft = 10;
        textBox.InternalMarginRight = 10;

        // Move the builder cursor inside the textbox.
        builder.MoveTo(textBoxShape.LastParagraph);

        // Apply bold formatting.
        builder.Font.Bold = true;

        // Insert a paragraph with bold text.
        builder.Writeln("This is bold text inside a textbox.");

        // Save the document.
        doc.Save("TextBoxWithMargins.docx");
    }
}
