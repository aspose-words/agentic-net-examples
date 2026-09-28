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
            // Create a new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a rectangle AutoShape with specific size.
            Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 200, 100);

            // Set the fill color to blue.
            rectangle.FillColor = Color.Blue;

            // Define the line (stroke) weight.
            rectangle.StrokeWeight = 2.0; // points
            rectangle.StrokeColor = Color.Black; // optional: set stroke color

            // Save the document.
            string outputPath = "RectangleShape.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
