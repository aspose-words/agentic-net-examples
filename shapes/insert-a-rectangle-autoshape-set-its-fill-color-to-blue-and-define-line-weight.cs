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

        // Insert a rectangle AutoShape with a width of 100 points and a height of 50 points.
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 100, 50);

        // Set the fill color of the shape to blue.
        rectangle.FillColor = Color.Blue;

        // Define the line (stroke) weight of the shape.
        rectangle.StrokeWeight = 2.0; // Thickness in points.

        // Save the document to a file in the current directory.
        string outputFile = Path.Combine(Directory.GetCurrentDirectory(), "RectangleShape.docx");
        doc.Save(outputFile);

        // Verify that the file was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException("The document was not saved correctly.");

        // Indicate successful completion.
        Console.WriteLine($"Document saved to: {outputFile}");
    }
}
