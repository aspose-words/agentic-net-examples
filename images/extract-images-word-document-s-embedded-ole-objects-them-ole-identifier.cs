using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractOleImages
{
    static void Main()
    {
        // Path to the source Word document.
        string inputPath = @"C:\Docs\InputDocument.docx";

        // Folder where extracted OLE objects will be saved.
        string outputFolder = @"C:\Docs\ExtractedOleObjects";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the Word document.
        Document doc = new Document(inputPath);

        // Get all Shape nodes in the document (including those inside headers/footers).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int oleIndex = 0;

        // Iterate through each shape and process embedded OLE objects.
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Access the OleFormat of the shape; null if the shape is not an OLE object.
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue; // Not an OLE object, skip.

            // Skip linked OLE objects; we only want embedded data.
            if (oleFormat.IsLink)
                continue;

            // Build a file name using the OLE object's ProgId (sanitized) and its suggested extension.
            // If the suggested extension is missing, default to a generic binary extension.
            string progIdSafe = oleFormat.ProgId.Replace(".", "_");
            string extension = string.IsNullOrEmpty(oleFormat.SuggestedExtension) ? ".bin" : oleFormat.SuggestedExtension;
            string fileName = Path.Combine(outputFolder,
                $"OleObject_{oleIndex}_{progIdSafe}{extension}");

            // Save the embedded OLE data directly to a file.
            oleFormat.Save(fileName);

            oleIndex++;
        }

        Console.WriteLine($"Extraction complete. {oleIndex} OLE objects saved to '{outputFolder}'.");
    }
}
