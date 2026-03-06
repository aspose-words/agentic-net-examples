using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading; // <-- added namespace for RtfLoadOptions

class ExtractOleObjectsFromRtf
{
    static void Main()
    {
        // Path to the source RTF document.
        string rtfPath = @"C:\Docs\SourceDocument.rtf";

        // Directory where extracted OLE objects will be saved.
        string outputDir = @"C:\Docs\ExtractedOleObjects";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputDir);

        // Load the RTF document. Use RtfLoadOptions to keep OLE data.
        Document doc = new Document(rtfPath, new RtfLoadOptions());

        // Get all shapes in the document (OLE objects are stored in shapes).
        var shapes = doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>();

        int oleIndex = 0;

        foreach (Shape shape in shapes)
        {
            // Access the OleFormat of the shape. If null, the shape does not contain OLE data.
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue;

            // Skip linked OLE objects – they cannot be saved directly.
            if (oleFormat.IsLink)
                continue;

            // Determine a suitable file extension for the extracted object.
            string extension = oleFormat.SuggestedExtension ?? ".bin";

            // Build the output file name.
            string outputPath = Path.Combine(outputDir, $"OleObject_{oleIndex}{extension}");

            // Save the OLE object to the file system.
            oleFormat.Save(outputPath);

            Console.WriteLine($"Extracted OLE object #{oleIndex} to: {outputPath}");

            oleIndex++;
        }

        Console.WriteLine("Extraction complete.");
    }
}
