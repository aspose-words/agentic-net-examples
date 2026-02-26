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
        Directory.CreateDirectory(outputDir);

        // Load the document.
        Document doc = new Document(inputPath);

        // Get all Shape nodes (OLE objects are stored in shapes).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int oleIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            // If the shape does not contain an OLE object, skip it.
            OleFormat ole = shape.OleFormat;
            if (ole == null)
                continue;

            // Determine a file name using the suggested extension.
            string extension = ole.SuggestedExtension ?? ".bin";
            string fileName = $"OleObject_{oleIndex}{extension}";
            string filePath = Path.Combine(outputDir, fileName);

            // Save the OLE object directly to a file.
            ole.Save(filePath);

            Console.WriteLine($"Extracted OLE object #{oleIndex} to: {filePath}");
            oleIndex++;
        }
    }
}
