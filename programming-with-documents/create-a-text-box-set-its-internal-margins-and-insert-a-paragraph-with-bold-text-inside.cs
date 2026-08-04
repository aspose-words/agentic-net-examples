using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text box shape with specific dimensions.
        Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 200, 100);

        // Access the TextBox object to set internal margins (in points).
        TextBox textBox = textBoxShape.TextBox;
        textBox.InternalMarginTop = 10;
        textBox.InternalMarginBottom = 10;
        textBox.InternalMarginLeft = 10;
        textBox.InternalMarginRight = 10;

        // Move the builder's cursor inside the text box.
        builder.MoveTo(textBoxShape.LastParagraph);

        // Set the font to bold.
        builder.Font.Bold = true;

        // Insert a paragraph with bold text.
        builder.Writeln("This is bold text inside a text box.");

        // Save the document to a file in the current directory.
        doc.Save("TextBoxMargins.docx");
    }
}
