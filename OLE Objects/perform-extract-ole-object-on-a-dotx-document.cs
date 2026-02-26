using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractOleFromDotx
{
    static void Main()
    {
        // Path to the DOTX template file
        string inputPath = @"C:\Docs\Template.dotx";

        // Directory where extracted OLE objects will be saved
        string outputDir = @"C:\Docs\ExtractedOle";
        Directory.CreateDirectory(outputDir);

        // Load the DOTX document
        Document doc = new Document(inputPath);

        // Get all shapes in the document (including those inside headers/footers)
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int oleCounter = 0;
        foreach (Shape shape in shapeNodes)
        {
            // Access the OLE format of the shape; null if not an OLE object
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue;

            // Skip linked OLE objects because they cannot be saved directly
            if (oleFormat.IsLink)
                continue;

            // Determine a suitable file extension; fallback to .bin if not provided
            string extension = oleFormat.SuggestedExtension ?? ".bin";

            // Build a unique file name for each extracted object
            string outputPath = Path.Combine(outputDir, $"OleObject_{oleCounter}{extension}");

            // Save the OLE object data to the file
            oleFormat.Save(outputPath);

            Console.WriteLine($"Extracted OLE object saved to: {outputPath}");
            oleCounter++;
        }

        Console.WriteLine("OLE extraction completed.");
    }
}
