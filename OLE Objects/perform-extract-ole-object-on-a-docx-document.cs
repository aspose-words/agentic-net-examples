using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractOleObjects
{
    static void Main()
    {
        // Path to the source DOCX file containing OLE objects.
        string docPath = @"C:\Docs\Sample.docx";

        // Directory where the extracted OLE files will be saved.
        string outputDir = @"C:\Docs\ExtractedOle";
        Directory.CreateDirectory(outputDir);

        // Load the Word document.
        Document doc = new Document(docPath);

        // Get all Shape nodes in the document (OLE objects are stored in shapes).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int oleIndex = 0;
        foreach (Shape shape in shapeNodes)
        {
            // Access the OleFormat of the shape; null if the shape is not an OLE object.
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue;

            // Skip linked OLE objects because they do not contain embedded data.
            if (oleFormat.IsLink)
                continue;

            // Determine a suitable file extension for the extracted object.
            string extension = oleFormat.SuggestedExtension ?? ".bin";

            // Build a unique file name for each extracted object.
            string outputPath = Path.Combine(outputDir, $"OleObject_{oleIndex}{extension}");

            // Save the embedded OLE data to the file.
            oleFormat.Save(outputPath);

            Console.WriteLine($"Extracted OLE object #{oleIndex} to: {outputPath}");
            oleIndex++;
        }
    }
}
