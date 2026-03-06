using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class OleExtractor
{
    static void Main()
    {
        // Path to the source DOCX file that contains OLE objects.
        string sourceDocPath = @"C:\Docs\SourceDocument.docx";

        // Folder where extracted OLE files will be saved.
        string outputFolder = @"C:\Docs\ExtractedOleObjects";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the Word document from the file system.
        Document doc = new Document(sourceDocPath);

        // Get all Shape nodes in the document (including those inside headers/footers).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int oleIndex = 0;

        // Iterate through each shape and process only OLE objects.
        foreach (Shape shape in shapeNodes)
        {
            if (shape.ShapeType != ShapeType.OleObject)
                continue; // Skip non‑OLE shapes.

            // Access the OleFormat of the shape.
            OleFormat oleFormat = shape.OleFormat;

            // Determine a suitable file extension for the embedded object.
            string extension = oleFormat.SuggestedExtension;

            // Build a unique file name for the extracted object.
            string outputFilePath = Path.Combine(outputFolder,
                $"OleObject_{oleIndex}{extension}");

            // Save the embedded OLE data directly to the file.
            oleFormat.Save(outputFilePath);

            Console.WriteLine($"Extracted OLE object #{oleIndex} to: {outputFilePath}");

            oleIndex++;
        }

        Console.WriteLine("Extraction complete.");
    }
}
