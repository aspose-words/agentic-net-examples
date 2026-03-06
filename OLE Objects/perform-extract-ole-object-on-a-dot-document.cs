using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractOleFromDot
{
    static void Main()
    {
        // Path to the DOT template file
        string inputPath = @"C:\Docs\Template.dot";

        // Directory where extracted OLE objects will be saved
        string outputDir = @"C:\Docs\ExtractedOle";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputDir);

        // Load the DOT document
        Document doc = new Document(inputPath);

        // Get all shapes in the document
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int oleIndex = 0;
        foreach (Shape shape in shapeNodes)
        {
            // Only process shapes that contain an OLE object
            OleFormat ole = shape.OleFormat;
            if (ole == null)
                continue;

            // Construct a file name using the suggested extension
            string fileName = $"OleObject_{oleIndex}{ole.SuggestedExtension}";
            string fullPath = Path.Combine(outputDir, fileName);

            // Save the OLE object directly to a file
            ole.Save(fullPath);

            oleIndex++;
        }
    }
}
