using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class OleExtractor
{
    static void Main()
    {
        // Path to the source DOCX file containing OLE objects.
        string inputPath = @"C:\Data\Input.docx";

        // Folder where extracted OLE files will be saved.
        string outputFolder = @"C:\Extracted\";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Get all Shape nodes (OLE objects are stored as shapes).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int oleIndex = 0;

        // Iterate through each shape and extract OLE data if present.
        foreach (Shape shape in shapeNodes)
        {
            // OleFormat is null for non‑OLE shapes.
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue;

            // Determine a suitable file extension for the embedded object.
            string extension = oleFormat.SuggestedExtension;

            // Build a unique file name for each extracted object.
            string outputPath = Path.Combine(outputFolder, $"OleObject_{oleIndex}{extension}");

            // Save the OLE object data directly to a file.
            oleFormat.Save(outputPath);

            oleIndex++;
        }
    }
}
