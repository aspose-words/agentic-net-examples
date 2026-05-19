using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Words.Lists;
using Aspose.Words.Saving;
using System.Linq;

public class UniformShapeFillExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a few sample shapes of different types.
        builder.InsertShape(ShapeType.Rectangle, 150, 80);
        builder.InsertShape(ShapeType.Ellipse, 100, 100);
        builder.InsertShape(ShapeType.Cloud, 120, 70);
        builder.InsertShape(ShapeType.Star, 80, 80);

        // Define the uniform fill color to be applied to every shape.
        Color brandColor = Color.FromArgb(0xFF, 0x4A, 0x90, 0xE2); // Example brand blue.

        // Traverse all shapes in the document and set their fill color.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Apply a solid fill color.
            shape.FillColor = brandColor;
        }

        // Save the modified document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "UniformFillShapes.docx");
        doc.Save(outputPath);

        // Validate that the file was created successfully.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");

        // The program ends here without waiting for user input.
    }
}
