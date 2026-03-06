using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractOleObjects
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = "Input.docx";

        // Directory where extracted OLE objects will be saved.
        string outputDir = "ExtractedOleObjects";
        Directory.CreateDirectory(outputDir);

        // Load the document.
        Document doc = new Document(inputPath);

        // Get all shapes in the document (including those inside headers/footers).
        var shapeNodes = doc.GetChildNodes(NodeType.Shape, true)
                            .OfType<Shape>();

        int oleIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            // Access the OLE data of the shape.
            OleFormat ole = shape.OleFormat;

            // Skip shapes that are not OLE objects or are linked (we only extract embedded data).
            if (ole == null || ole.IsLink)
                continue;

            // Determine a suitable file extension; fall back to .bin if none is suggested.
            string extension = ole.SuggestedExtension;
            if (string.IsNullOrEmpty(extension))
                extension = ".bin";

            // Build a unique file name for each extracted object.
            string filePath = Path.Combine(outputDir, $"OleObject_{oleIndex}{extension}");

            // Save the embedded OLE object directly to the file.
            ole.Save(filePath);

            oleIndex++;
        }

        Console.WriteLine($"Extraction complete. {oleIndex} OLE object(s) saved to '{outputDir}'.");
    }
}
