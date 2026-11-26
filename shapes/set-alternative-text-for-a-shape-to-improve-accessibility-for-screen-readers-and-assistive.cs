using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace ShapeAlternativeTextExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a simple rectangle shape.
            Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);

            // Set alternative text for accessibility.
            shape.AlternativeText = "Blue rectangle representing a placeholder image";

            // Save the document.
            string outputPath = "ShapeWithAltText.docx";
            doc.Save(outputPath);

            // Validate that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");

            // Optional: Verify that the shape's alternative text was set correctly.
            Shape savedShape = (Shape)doc.GetChildNodes(NodeType.Shape, true)[0];
            if (savedShape.AlternativeText != "Blue rectangle representing a placeholder image")
                throw new InvalidOperationException("Alternative text was not set correctly on the shape.");
        }
    }
}
