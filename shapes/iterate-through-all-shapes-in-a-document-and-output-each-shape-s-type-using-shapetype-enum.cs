using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a few sample shapes into the document.
        builder.InsertShape(ShapeType.Rectangle, 100, 50);
        builder.InsertShape(ShapeType.Ellipse, 80, 80);
        builder.InsertShape(ShapeType.Star, 60, 60);

        // Save the document (optional, just to have an output file).
        doc.Save("ShapesOutput.docx");

        // Retrieve all shape nodes in the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate through each shape and write its ShapeType to the console.
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            Console.WriteLine(shape.ShapeType);
        }
    }
}
