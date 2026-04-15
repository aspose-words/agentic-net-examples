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

        // Insert a floating text box shape.
        Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 200, 100);
        // Add some text inside the text box.
        builder.MoveTo(textBoxShape.LastParagraph);
        builder.Writeln("Sample text inside the text box.");

        // Retrieve the parent shape of the text box via the TextBox property.
        Shape parentShape = textBoxShape.TextBox.Parent;
        // Change the rotation angle of the parent shape to 45 degrees.
        parentShape.Rotation = 45;

        // Define an output path relative to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TextBoxParentRotated.docx");
        // Save the modified document.
        doc.Save(outputPath);
    }
}
