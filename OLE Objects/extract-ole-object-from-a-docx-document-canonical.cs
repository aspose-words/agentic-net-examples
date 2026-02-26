using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class OleExtractor
{
    static void Main()
    {
        // Path to the source DOCX document that contains OLE objects.
        string sourceDocPath = @"C:\Docs\input.docx";

        // Folder where extracted OLE files will be saved.
        string outputFolder = @"C:\Docs\ExtractedOle";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the Word document.
        Document doc = new Document(sourceDocPath);

        // Get all Shape nodes in the document (including those inside headers/footers).
        var shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate through each shape and extract OLE objects.
        foreach (Shape shape in shapeNodes)
        {
            // Only shapes that actually contain an OLE object have a non‑null OleFormat.
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue;

            // Determine a suitable file name for the extracted object.
            // Use the suggested file extension for the embedded object.
            string extension = oleFormat.SuggestedExtension ?? ".bin";

            // Create a unique file name based on the shape index.
            string fileName = Path.Combine(outputFolder,
                $"OleObject_{shape.GetHashCode()}{extension}");

            // Save the embedded OLE data directly to a file.
            oleFormat.Save(fileName);

            Console.WriteLine($"Extracted OLE object saved to: {fileName}");
        }
    }
}
