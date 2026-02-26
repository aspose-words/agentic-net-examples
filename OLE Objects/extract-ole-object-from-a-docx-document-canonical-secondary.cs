using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class OleExtractor
{
    static void Main()
    {
        // Path to the source DOCX document.
        string sourcePath = "InputDocument.docx";

        // Directory where extracted OLE objects will be saved.
        string outputFolder = "ExtractedOleObjects";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the DOCX document.
        Document doc = new Document(sourcePath);

        // Get all shape nodes in the document (including those inside headers/footers).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int oleIndex = 0;

        // Iterate through each shape and process OLE objects.
        foreach (Shape shape in shapeNodes)
        {
            // Only shapes that contain an OLE object have a non‑null OleFormat.
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue;

            // Skip linked OLE objects; they do not contain embedded data to extract.
            if (oleFormat.IsLink)
                continue;

            // Determine a suitable file extension for the extracted object.
            // If the document does not suggest one, default to a generic binary extension.
            string extension = oleFormat.SuggestedExtension ?? ".bin";

            // Build the output file name.
            string outputPath = Path.Combine(outputFolder, $"OleObject_{oleIndex}{extension}");

            // Save the embedded OLE object to the file system.
            oleFormat.Save(outputPath);

            Console.WriteLine($"Extracted OLE object #{oleIndex} to \"{outputPath}\"");
            oleIndex++;
        }

        Console.WriteLine("Extraction complete.");
    }
}
