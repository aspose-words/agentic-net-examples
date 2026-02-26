using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractOleObjects
{
    static void Main()
    {
        // Path to the DOCX file that contains OLE objects.
        string inputPath = @"C:\Docs\Sample.docx";

        // Directory where the extracted OLE files will be saved.
        string outputDir = @"C:\Docs\ExtractedOle";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputDir);

        // Load the document.
        Document doc = new Document(inputPath);

        // Get all Shape nodes (they may contain OLE objects).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int oleIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            // Skip shapes that are not OLE objects.
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue;

            // Determine a suitable file extension; fall back to .bin if unknown.
            string extension = oleFormat.SuggestedExtension ?? ".bin";

            // Build a unique file name for each extracted object.
            string outputPath = Path.Combine(outputDir, $"OleObject_{oleIndex}{extension}");

            // Save the OLE object directly to the file system.
            oleFormat.Save(outputPath);

            Console.WriteLine($"Extracted OLE object saved to: {outputPath}");
            oleIndex++;
        }
    }
}
