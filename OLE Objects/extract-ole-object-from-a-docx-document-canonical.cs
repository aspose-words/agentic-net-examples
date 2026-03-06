using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractOleObjects
{
    static void Main()
    {
        // Path to the DOCX file that contains OLE objects.
        string inputPath = @"C:\Docs\SampleWithOle.docx";

        // Directory where the extracted OLE files will be saved.
        string outputDir = @"C:\Docs\ExtractedOle";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputDir);

        // Load the Word document.
        Document doc = new Document(inputPath);

        // Get all Shape nodes (they may contain OLE objects).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int oleIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            // Access the OleFormat of the shape; null if the shape is not an OLE object.
            OleFormat ole = shape.OleFormat;
            if (ole == null)
                continue;

            // Determine a file name for the extracted object.
            string extension = ole.SuggestedExtension ?? ".bin";
            string outFile = Path.Combine(outputDir, $"OleObject_{oleIndex}{extension}");

            // Save the embedded OLE object directly to a file.
            ole.Save(outFile);

            oleIndex++;
        }
    }
}
