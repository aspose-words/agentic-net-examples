using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractOleObjects
{
    static void Main()
    {
        // Path to the source DOTM document
        string inputPath = @"C:\Docs\Template.dotm";

        // Directory where extracted OLE objects will be saved
        string outputDir = @"C:\Docs\ExtractedOle";
        Directory.CreateDirectory(outputDir);

        // Load the DOTM document
        Document doc = new Document(inputPath);

        // Get all shapes in the document (including those inside headers/footers)
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int oleIndex = 0;
        foreach (Shape shape in shapeNodes)
        {
            // Only process shapes that contain an OLE object
            OleFormat ole = shape.OleFormat;
            if (ole == null)
                continue;

            // Determine a file name using the suggested extension (fallback to .bin)
            string extension = ole.SuggestedExtension ?? ".bin";
            string outputPath = Path.Combine(outputDir, $"OleObject_{oleIndex}{extension}");

            // Save the OLE object data to the file system
            ole.Save(outputPath);

            oleIndex++;
        }
    }
}
