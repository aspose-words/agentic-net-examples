using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with some text that will appear over the shape.
        builder.Writeln("This text should be visible over the shape after it is sent to the back.");

        // Insert a rectangle shape that overlaps the text.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 200, 100);
        shape.FillColor = Color.LightBlue;
        shape.StrokeColor = Color.DarkBlue;
        shape.WrapType = WrapType.None; // Ensure the shape can overlap text.

        // Send the shape to the back of the layering order.
        shape.BehindText = true;

        // Save the document.
        string outputPath = "ShapeBack.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Optionally, clean up (not required for the task).
    }
}
