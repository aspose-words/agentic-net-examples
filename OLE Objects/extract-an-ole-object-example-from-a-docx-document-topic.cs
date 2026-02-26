using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Path to the folder that contains the source DOCX and where extracted files will be saved.
        string dataDir = @"C:\Data\";

        // Load the DOCX document that contains OLE objects.
        Document doc = new Document(Path.Combine(dataDir, "Input.docx"));

        // Get all Shape nodes in the document (including those inside headers/footers).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;

        // Iterate through each shape and process only OLE objects.
        foreach (Shape shape in shapeNodes)
        {
            if (shape.ShapeType != ShapeType.OleObject)
                continue;

            // Access the OleFormat of the shape.
            OleFormat oleFormat = shape.OleFormat;

            // Determine a suitable file extension for the embedded object.
            string extension = oleFormat.SuggestedExtension ?? ".bin";

            // Build the output file name.
            string outFile = Path.Combine(dataDir, $"Extracted_{extractedCount}{extension}");

            // Save the embedded OLE object directly to a file.
            oleFormat.Save(outFile);

            extractedCount++;
        }

        // Optional: inform how many objects were extracted.
        Console.WriteLine($"{extractedCount} OLE object(s) extracted.");
    }
}
