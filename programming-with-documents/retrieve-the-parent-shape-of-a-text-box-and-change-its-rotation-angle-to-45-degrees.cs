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
        builder.Writeln("Sample text inside the textbox.");

        // Retrieve the parent shape of the TextBox.
        Shape parentShape = textBoxShape.TextBox.Parent;

        // Set the rotation angle of the parent shape to 45 degrees.
        parentShape.Rotation = 45;

        // Ensure the output directory exists.
        string artifactsDir = "Artifacts";
        Directory.CreateDirectory(artifactsDir);

        // Save the modified document.
        doc.Save(Path.Combine(artifactsDir, "TextBoxRotated.docx"));
    }
}
