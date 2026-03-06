using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class OleExtractor
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourceDocPath = @"C:\Docs\SourceDocument.docx";

        // Directory where extracted OLE objects will be saved.
        string outputDirectory = @"C:\Docs\ExtractedOleObjects";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputDirectory);

        // Load the DOCX document.
        Document doc = new Document(sourceDocPath);

        // Retrieve all shape nodes in the document.
        var shapes = doc.GetChildNodes(NodeType.Shape, true)
                        .OfType<Shape>()
                        .Where(s => s.ShapeType == ShapeType.OleObject);

        int index = 1;
        foreach (Shape shape in shapes)
        {
            // Access the OLE format of the shape.
            OleFormat oleFormat = shape.OleFormat;

            // Skip linked OLE objects; only extract embedded ones.
            if (oleFormat.IsLink)
                continue;

            // Determine a suitable file extension for the embedded object.
            string extension = oleFormat.SuggestedExtension ?? ".bin";

            // Build the output file name.
            string outputPath = Path.Combine(outputDirectory,
                                             $"OleObject_{index}{extension}");

            // Save the embedded OLE object to the file system.
            oleFormat.Save(outputPath);

            Console.WriteLine($"Extracted OLE object saved to: {outputPath}");
            index++;
        }

        Console.WriteLine("Extraction complete.");
    }
}
