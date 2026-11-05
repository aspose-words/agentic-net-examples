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

        // Define shape dimensions (points).
        double shapeWidth = 200;
        double shapeHeight = 100;

        // Insert a rectangle shape. Rounded corners are not directly supported in this version,
        // so we use a regular rectangle and apply the required fill and border.
        Shape rectShape = builder.InsertShape(ShapeType.Rectangle, shapeWidth, shapeHeight);

        // Set a light gray fill.
        rectShape.FillColor = Color.LightGray;

        // Optional: set a thin black border.
        rectShape.StrokeColor = Color.Black;
        rectShape.StrokeWeight = 0.5;

        // Save the document.
        string outputPath = "CustomSnipCornerRectangle.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
