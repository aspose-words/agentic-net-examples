using System;
using System.Drawing;
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

        // Insert a rectangle AutoShape with specific size.
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 150, 80);

        // Set the fill color to blue.
        rectangle.FillColor = Color.Blue;

        // Define the line (stroke) weight.
        rectangle.StrokeWeight = 2.0; // points

        // Save the document.
        string outputPath = "RectangleShape.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved successfully.");

        // Optional: indicate success (no user interaction required).
        Console.WriteLine("Document created: " + Path.GetFullPath(outputPath));
    }
}
