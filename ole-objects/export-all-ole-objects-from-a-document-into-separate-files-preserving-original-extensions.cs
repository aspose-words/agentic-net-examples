using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Path to the source Word document containing OLE objects.
        // The file must exist in the working directory; otherwise the program will exit gracefully.
        string inputPath = "InputDocument.docx";

        // Directory where extracted OLE files will be saved.
        string outputDir = "ExtractedOleObjects";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputDir);

        // Verify that the input document exists before attempting to load it.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found. No OLE objects were extracted.");
            return;
        }

        // Load the Word document.
        Document doc = new Document(inputPath);

        // Get all shapes in the document (including those inside headers/footers).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        int oleIndex = 0;
        foreach (Shape shape in shapes)
        {
            // Only process shapes that contain OLE data.
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue;

            // Skip linked OLE objects because they cannot be saved directly.
            if (oleFormat.IsLink)
                continue;

            // Determine a file name for the extracted object.
            // Use the suggested extension if available; otherwise default to ".bin".
            string extension = oleFormat.SuggestedExtension ?? ".bin";
            string fileName = $"OleObject_{oleIndex}{extension}";
            string fullPath = Path.Combine(outputDir, fileName);

            // Save the OLE object to the file system.
            oleFormat.Save(fullPath);
            Console.WriteLine($"Saved OLE object to: {fullPath}");

            oleIndex++;
        }

        if (oleIndex == 0)
        {
            Console.WriteLine("No embedded OLE objects were found in the document.");
        }
    }
}
