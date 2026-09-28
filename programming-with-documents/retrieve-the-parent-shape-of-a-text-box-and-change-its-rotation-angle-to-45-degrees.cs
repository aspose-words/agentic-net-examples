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

        // Insert a text box shape.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);

        // Rotate the shape (its parent) to 45 degrees.
        textBox.Rotation = 45;

        // Save the document.
        doc.Save("Output.docx");
    }
}
