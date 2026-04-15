using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a textbox shape with specific size.
        Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 200, 100);

        // Access the TextBox object to set internal margins (in points).
        TextBox textBox = textBoxShape.TextBox;
        textBox.InternalMarginTop = 15;
        textBox.InternalMarginBottom = 15;
        textBox.InternalMarginLeft = 15;
        textBox.InternalMarginRight = 15;

        // Move the builder cursor inside the textbox.
        builder.MoveTo(textBoxShape.LastParagraph);

        // Set the font to bold and write a paragraph.
        builder.Font.Bold = true;
        builder.Writeln("Bold text inside the textbox.");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TextBoxMargins.docx");
        doc.Save(outputPath);
    }
}
