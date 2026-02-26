using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractOleObjects
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("Input.doc");

        // Get all shapes in the document (including those inside headers/footers).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int oleIndex = 0;
        foreach (Shape shape in shapeNodes)
        {
            // Only process shapes that actually contain an OLE object.
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue;

            // Determine a file name for the extracted object.
            // Use the suggested extension; fall back to .bin if none is provided.
            string extension = oleFormat.SuggestedExtension ?? ".bin";
            string outputPath = $"OleObject_{oleIndex}{extension}";

            // Save the embedded OLE data to the file system.
            oleFormat.Save(outputPath);

            Console.WriteLine($"Extracted OLE object #{oleIndex} to \"{outputPath}\"");
            oleIndex++;
        }
    }
}
