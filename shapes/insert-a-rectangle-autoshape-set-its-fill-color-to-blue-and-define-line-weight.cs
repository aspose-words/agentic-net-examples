using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace ShapeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a rectangle AutoShape with specified width and height (in points).
            Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 200, 100);

            // Set the fill color of the rectangle to blue.
            rectangle.FillColor = Color.Blue;

            // Define the line (stroke) weight of the rectangle.
            rectangle.StrokeWeight = 2.0; // Weight in points.

            // Save the document to the local file system.
            string outputPath = "RectangleShape.docx";
            doc.Save(outputPath);

            // Validate that the file was created successfully.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
