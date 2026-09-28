using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Define output file path
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CustomShapeMetadata.docx");

        // Create a new blank document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape and set its size
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 150, 80);
        // Store custom metadata in the Title property
        shape.Title = "CustomMetadata:SampleValue";

        // Save the document
        doc.Save(outputPath);

        // Validate that the file was created
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");

        // Load the document for processing
        Document loadedDoc = new Document(outputPath);

        // Traverse all shapes in the document
        var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                              .Cast<Shape>();

        // Retrieve the custom metadata from the Title property
        foreach (var s in shapes)
        {
            if (!string.IsNullOrEmpty(s.Title) && s.Title.StartsWith("CustomMetadata:"))
            {
                string metadata = s.Title.Substring("CustomMetadata:".Length);
                Console.WriteLine($"Retrieved metadata: {metadata}");
                // Since we only expect one shape, we can break after finding it
                break;
            }
        }
    }
}
