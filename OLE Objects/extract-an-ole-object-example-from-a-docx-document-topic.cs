using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class ExtractOleObjects
{
    static void Main()
    {
        // Path to the source DOCX document that contains OLE objects.
        string inputPath = @"C:\Docs\DocumentWithOle.docx";

        // Folder where extracted OLE files will be saved.
        string outputFolder = @"C:\Docs\ExtractedOle";
        Directory.CreateDirectory(outputFolder);

        // Load the document using the Document(string) constructor (lifecycle rule).
        Document doc = new Document(inputPath);

        // Get all Shape nodes in the document (including those inside headers/footers).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int oleIndex = 0;
        foreach (Shape shape in shapeNodes)
        {
            // Process only shapes that actually contain an OLE object.
            if (shape.ShapeType != ShapeType.OleObject)
                continue;

            OleFormat oleFormat = shape.OleFormat;

            // Determine a suitable file extension for the embedded object.
            string extension = oleFormat.SuggestedExtension ?? ".bin";

            // Build a unique file name for each extracted object.
            string extractedPath = Path.Combine(outputFolder,
                $"OleObject_{oleIndex}{extension}");

            // Save the OLE object data directly to a file (uses OleFormat.Save(string) rule).
            oleFormat.Save(extractedPath);

            Console.WriteLine($"Extracted OLE object #{oleIndex} to: {extractedPath}");
            oleIndex++;
        }

        // Optional: inform if no OLE objects were found.
        if (oleIndex == 0)
            Console.WriteLine("No OLE objects were found in the document.");
    }
}
