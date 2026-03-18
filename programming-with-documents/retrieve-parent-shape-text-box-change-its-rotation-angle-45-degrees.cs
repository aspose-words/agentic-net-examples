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

        // Insert a textbox shape into the document.
        Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 200, 100);

        // Add some text inside the textbox.
        builder.MoveTo(textBoxShape.LastParagraph);
        builder.Writeln("Sample text inside the textbox.");

        // Retrieve the parent shape of the textbox.
        Shape parentShape = textBoxShape.TextBox.Parent;

        // Rotate the parent shape 45 degrees clockwise.
        parentShape.Rotation = 45;

        // Save the modified document.
        doc.Save("ParentShapeRotated.docx");
    }
}
