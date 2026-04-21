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

        // Insert a rectangle shape with an initial size.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 200, 100);

        // Lock the aspect ratio of the shape.
        shape.AspectRatioLocked = true;

        // Define the output file name.
        string outputPath = "AspectRatioLocked.docx";

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"The output file '{outputPath}' was not created.");
        }

        // Optionally, you could add further validation here (e.g., re-open and check the property),
        // but the task only requires setting the property and saving the document.
    }
}
