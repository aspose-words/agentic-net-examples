using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class ShapeIterationExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a few sample shapes using the builder.
        builder.InsertShape(ShapeType.Rectangle, 100.0, 50.0);
        builder.InsertShape(ShapeType.Ellipse, 80.0, 80.0);
        builder.InsertShape(ShapeType.Star, 60.0, 60.0);

        // Define the output file path and save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleShapes.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to save the document.");

        // Retrieve all shape nodes from the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate through each shape and output its ShapeType.
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            Console.WriteLine(shape.ShapeType);
        }
    }
}
