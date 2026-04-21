using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class UniformShapeFillExample
{
    public static void Main()
    {
        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "UniformFillShapes.docx");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a few sample shapes using DocumentBuilder.InsertShape.
        builder.InsertShape(ShapeType.Rectangle, 120, 60);
        builder.InsertShape(ShapeType.Ellipse, 80, 80);
        builder.InsertShape(ShapeType.CloudCallout, 150, 100);

        // Traverse all shapes in the document and apply a uniform fill color.
        var shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Set the fill color to a branding color (e.g., CornflowerBlue).
            shape.FillColor = System.Drawing.Color.CornflowerBlue;
        }

        // Save the modified document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file at '{outputPath}'.");
        }

        // Optionally, inform that the process completed successfully.
        Console.WriteLine($"Document saved successfully to: {outputPath}");
    }
}
