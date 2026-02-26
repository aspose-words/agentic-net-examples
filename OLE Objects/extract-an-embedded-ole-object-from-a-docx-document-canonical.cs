using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class OleExtractor
{
    static void Main()
    {
        // Path to the source DOCX file that contains embedded OLE objects.
        string sourceDocPath = @"C:\Docs\SourceDocument.docx";

        // Directory where the extracted OLE files will be saved.
        string outputDir = @"C:\Docs\ExtractedOleObjects";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputDir);

        // Load the Word document.
        Document doc = new Document(sourceDocPath);

        // Retrieve all Shape nodes in the document (OLE objects are stored in Shape nodes).
        var oleShapes = doc.GetChildNodes(NodeType.Shape, true)
                           .Cast<Shape>()
                           .Where(s => s.OleFormat != null);

        int index = 0;
        foreach (Shape shape in oleShapes)
        {
            OleFormat oleFormat = shape.OleFormat;

            // Skip linked OLE objects; we only extract embedded ones.
            if (oleFormat.IsLink)
                continue;

            // Determine a suitable file extension for the extracted object.
            string extension = oleFormat.SuggestedExtension ?? ".bin";

            // Build the output file name.
            string outputFilePath = Path.Combine(outputDir,
                $"ExtractedObject_{index}{extension}");

            // Save the OLE object data directly to a file.
            oleFormat.Save(outputFilePath);

            Console.WriteLine($"Extracted OLE object saved to: {outputFilePath}");
            index++;
        }

        Console.WriteLine("Extraction completed.");
    }
}
