using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class OleExtractor
{
    static void Main()
    {
        // Path to the source DOC document.
        string sourceDocPath = @"C:\Docs\sample.doc";

        // Folder where extracted OLE objects will be saved.
        string outputFolder = @"C:\ExtractedOle";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the document.
        Document doc = new Document(sourceDocPath);

        // Get all Shape nodes (they may contain OLE objects).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int oleCounter = 0;

        foreach (Shape shape in shapeNodes)
        {
            // Access the OleFormat of the shape; null if the shape is not an OLE object.
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue;

            // Skip linked OLE objects – saving them throws InvalidOperationException.
            if (oleFormat.IsLink)
                continue;

            // Determine a file name using the suggested extension (e.g., ".xlsx").
            string extension = oleFormat.SuggestedExtension ?? string.Empty;
            string fileName = $"OleObject_{oleCounter}{extension}";
            string fullPath = Path.Combine(outputFolder, fileName);

            // Save the embedded OLE object directly to a file.
            oleFormat.Save(fullPath);

            Console.WriteLine($"Extracted OLE object saved to: {fullPath}");
            oleCounter++;
        }
    }
}
