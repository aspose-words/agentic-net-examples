using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class OleExtractor
{
    static void Main()
    {
        // Path to the source DOCX file that contains OLE objects.
        string sourceDocPath = @"C:\Input\DocumentWithOle.docx";

        // Folder where extracted OLE files will be saved.
        string outputFolder = @"C:\Output\ExtractedOle";
        Directory.CreateDirectory(outputFolder);

        // Load the Word document.
        Document doc = new Document(sourceDocPath);

        // Get all Shape nodes in the document (OLE objects are stored in Shape nodes).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int oleIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            // Only process shapes that actually contain an OLE object.
            if (shape.ShapeType == ShapeType.OleObject && shape.OleFormat != null)
            {
                OleFormat oleFormat = shape.OleFormat;

                // Determine a suitable file name for the extracted object.
                string extension = oleFormat.SuggestedExtension ?? ".bin";
                string fileName = $"OleObject_{oleIndex}{extension}";
                string outputPath = Path.Combine(outputFolder, fileName);

                // Save the OLE object data directly to a file.
                oleFormat.Save(outputPath);

                Console.WriteLine($"Extracted OLE object #{oleIndex} to: {outputPath}");
                oleIndex++;
            }
        }

        Console.WriteLine("Extraction complete.");
    }
}
