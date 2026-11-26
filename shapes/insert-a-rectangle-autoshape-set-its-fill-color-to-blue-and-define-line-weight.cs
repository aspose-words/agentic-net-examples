using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace ShapeExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a rectangle AutoShape with specific size.
            Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 150, 100);

            // Set the fill color to blue.
            rectangle.FillColor = Color.Blue;

            // Define the line (stroke) weight.
            rectangle.StrokeWeight = 2.0; // points

            // Save the document.
            string outputPath = "RectangleShape.docx";
            doc.Save(outputPath);

            // Validate that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
