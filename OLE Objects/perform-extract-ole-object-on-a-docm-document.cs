using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractOleObjects
{
    static void Main()
    {
        // Path to the DOCM file that contains OLE objects.
        string inputFile = @"C:\Docs\OLE spreadsheet.docm";

        // Directory where extracted OLE files will be saved.
        string outputDir = @"C:\ExtractedOle\";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputDir);

        // Load the document (uses the provided Document constructor rule).
        Document doc = new Document(inputFile);

        // Get all shape nodes in the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int oleIndex = 0;

        // Iterate through each shape and extract OLE objects.
        foreach (Shape shape in shapeNodes)
        {
            // Only shapes that contain an OLE object have a non‑null OleFormat.
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue;

            // Determine a suitable file extension for the extracted object.
            string extension = oleFormat.SuggestedExtension ?? string.Empty;

            // Build the output file name.
            string outputPath = Path.Combine(outputDir, $"OleObject_{oleIndex}{extension}");

            // Save the OLE object data to a file (uses OleFormat.Save(string) rule).
            oleFormat.Save(outputPath);

            Console.WriteLine($"Extracted OLE object #{oleIndex} to: {outputPath}");
            oleIndex++;
        }

        Console.WriteLine("Extraction complete.");
    }
}
