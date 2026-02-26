using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the HTML document.
        Document doc = new Document("input.html");

        // Find all shapes that contain OLE objects.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int index = 0;

        foreach (Shape shape in shapeNodes)
        {
            // Only process shapes that actually have an OLE object.
            OleFormat ole = shape.OleFormat;
            if (ole == null)
                continue;

            // Determine a suitable file extension for the extracted object.
            string extension = ole.SuggestedExtension ?? string.Empty;

            // Build a unique file name for each extracted OLE object.
            string outputPath = $"ExtractedObject_{index}{extension}";

            // Save the OLE object data directly to a file.
            ole.Save(outputPath);

            Console.WriteLine($"Extracted OLE object saved to: {outputPath}");
            index++;
        }
    }
}
