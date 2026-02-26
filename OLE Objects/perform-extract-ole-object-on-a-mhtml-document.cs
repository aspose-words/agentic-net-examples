using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractOleFromMhtml
{
    static void Main()
    {
        // Path to the source MHTML document
        string inputPath = @"C:\Docs\source.mht";

        // Directory where extracted OLE files will be saved
        string outputDir = @"C:\Docs\ExtractedOle";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputDir);

        // Load the MHTML document
        Document doc = new Document(inputPath);

        // Get all Shape nodes (OLE objects are stored in shapes)
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int oleIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            // Access the OleFormat of the shape; null if the shape is not an OLE object
            OleFormat ole = shape.OleFormat;
            if (ole == null)
                continue;

            // Skip linked OLE objects – they cannot be saved directly
            if (ole.IsLink)
                continue;

            // Determine a suitable file extension for the embedded object
            string extension = ole.SuggestedExtension ?? ".bin";

            // Build a unique file name for each extracted OLE object
            string outputPath = Path.Combine(outputDir, $"OleObject_{oleIndex}{extension}");

            // Save the OLE object data to the file
            ole.Save(outputPath);

            Console.WriteLine($"Extracted OLE object saved to: {outputPath}");
            oleIndex++;
        }
    }
}
