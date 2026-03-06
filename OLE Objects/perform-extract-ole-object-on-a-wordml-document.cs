using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class OleExtractor
{
    static void Main()
    {
        // Path to the source WORDML (or any Word) document.
        string sourcePath = @"C:\Docs\SourceDocument.docx";

        // Directory where extracted OLE objects will be saved.
        string outputDir = @"C:\Docs\ExtractedOleObjects";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputDir);

        // Load the document.
        Document doc = new Document(sourcePath);

        // Find all shapes that contain OLE objects.
        var oleShapes = doc.GetChildNodes(NodeType.Shape, true)
                           .OfType<Shape>()
                           .Where(s => s.OleFormat != null);

        int index = 0;
        foreach (Shape shape in oleShapes)
        {
            OleFormat ole = shape.OleFormat;

            // Skip linked OLE objects – they have no embedded data to extract.
            if (ole.IsLink)
                continue;

            // Determine a suitable file extension for the extracted object.
            string extension = ole.SuggestedExtension ?? ".bin";

            // Build a unique file name for each extracted object.
            string filePath = Path.Combine(outputDir, $"OleObject_{index}{extension}");

            // Save the OLE object data directly to a file.
            ole.Save(filePath);

            Console.WriteLine($"Extracted OLE object #{index} to: {filePath}");
            index++;
        }

        Console.WriteLine("Extraction complete.");
    }
}
