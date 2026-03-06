using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class OleExtractor
{
    static void Main()
    {
        // Path to the source DOCX document that contains OLE objects.
        string inputPath = @"MyDir\DocumentWithOle.docx";

        // Directory where extracted OLE files will be saved.
        string outputDir = @"ArtifactsDir\ExtractedOleObjects";
        Directory.CreateDirectory(outputDir);

        // Load the document.
        Document doc = new Document(inputPath);

        // Get all Shape nodes in the document (including those inside headers/footers).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int oleIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            // Only process shapes that actually contain an OLE object.
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue;

            // Determine a file name for the extracted OLE object.
            // Use the suggested extension (e.g., ".xlsx") and a sequential index.
            string fileName = $"OleObject_{oleIndex}{oleFormat.SuggestedExtension}";
            string outputPath = Path.Combine(outputDir, fileName);

            // Save the OLE object data directly to a file.
            // This works for embedded objects; linked objects will throw InvalidOperationException.
            oleFormat.Save(outputPath);

            Console.WriteLine($"Extracted OLE object #{oleIndex} to: {outputPath}");
            oleIndex++;
        }

        Console.WriteLine("Extraction complete.");
    }
}
