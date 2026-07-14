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

        // Insert a text box shape into the document.
        Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 200, 100);
        // Add some sample text inside the text box.
        builder.MoveTo(textBoxShape.LastParagraph);
        builder.Writeln("Sample text inside the text box.");

        // Retrieve the parent shape of the TextBox (which is the shape itself).
        Shape parentShape = textBoxShape.TextBox.Parent;

        // Change the rotation angle of the parent shape to 45 degrees.
        parentShape.Rotation = 45;

        // Save the document to a file.
        string outputPath = "ParentShapeRotated.docx";
        doc.Save(outputPath);
    }
}
