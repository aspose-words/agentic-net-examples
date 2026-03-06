using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;

class ExtractOleFromTxt
{
    static void Main()
    {
        // Path to the source TXT document that may contain embedded OLE objects.
        string inputPath = @"C:\Docs\source.txt";

        // Directory where extracted OLE objects will be saved.
        string outputDir = @"C:\Docs\ExtractedOle";
        Directory.CreateDirectory(outputDir);

        // Load the TXT document with default load options.
        Document doc = new Document(inputPath, new TxtLoadOptions());

        // Iterate through all shapes in the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int oleIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            // Process only shapes that are OLE objects.
            if (shape.ShapeType != ShapeType.OleObject)
                continue;

            OleFormat ole = shape.OleFormat;

            // Skip linked OLE objects; only extract embedded ones.
            if (ole.IsLink)
                continue;

            // Determine a suitable file extension for the extracted object.
            string extension = ole.SuggestedExtension ?? ".bin";

            // Build the output file name.
            string outFile = Path.Combine(outputDir, $"OleObject_{oleIndex}{extension}");

            // Save the OLE object data to the file.
            ole.Save(outFile);

            oleIndex++;
        }

        Console.WriteLine($"Extraction complete. {oleIndex} OLE object(s) saved to '{outputDir}'.");
    }
}
