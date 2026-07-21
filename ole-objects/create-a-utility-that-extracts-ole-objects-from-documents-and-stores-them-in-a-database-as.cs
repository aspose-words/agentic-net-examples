using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class OleExtractor
{
    public static void Main()
    {
        // Path to the source Word document containing OLE objects.
        const string sourceDocPath = "SampleDocument.docx";

        // Verify that the source document exists.
        if (!File.Exists(sourceDocPath))
        {
            Console.WriteLine($"Source document not found: {Path.GetFullPath(sourceDocPath)}");
            return;
        }

        // Directory where extracted OLE objects will be saved.
        const string outputDirectory = "ExtractedOleObjects";

        // Load the Word document.
        Document doc = new Document(sourceDocPath);

        // Ensure the output directory exists.
        if (!Directory.Exists(outputDirectory))
            Directory.CreateDirectory(outputDirectory);

        // Iterate over all shapes in the document.
        int oleIndex = 0;
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue; // Not an OLE object.

            // Skip linked OLE objects; only process embedded ones.
            if (oleFormat.IsLink)
                continue;

            // Determine a file name for the object.
            string fileName = oleFormat.SuggestedFileName;
            if (string.IsNullOrEmpty(fileName))
            {
                // Fallback to a generated name using the suggested extension.
                string extension = oleFormat.SuggestedExtension ?? ".bin";
                fileName = $"OleObject_{oleIndex}{extension}";
            }

            // Build the full path for the output file.
            string outputPath = Path.Combine(outputDirectory, fileName);

            // Save the OLE object directly to the file.
            oleFormat.Save(outputPath);

            Console.WriteLine($"Extracted OLE object saved to: {outputPath}");
            oleIndex++;
        }

        Console.WriteLine("OLE extraction completed.");
    }
}
