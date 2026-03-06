using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractOleObjectsFromDotx
{
    static void Main()
    {
        // Path to the source DOTX template.
        string sourcePath = @"C:\Input\Template.dotx";

        // Folder where extracted OLE objects will be saved.
        string outputFolder = @"C:\Output\OleObjects";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the DOTX document.
        Document doc = new Document(sourcePath);

        // Get a flat list of all Shape nodes in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate through the shapes using an index so we can build unique file names.
        for (int i = 0; i < shapes.Count; i++)
        {
            Shape shape = (Shape)shapes[i];

            // Check if the shape contains an OLE object.
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue; // Not an OLE object, skip.

            // Determine a file name for the extracted object.
            // Use the suggested extension if available; otherwise default to .bin.
            string extension = string.IsNullOrEmpty(oleFormat.SuggestedExtension)
                ? ".bin"
                : oleFormat.SuggestedExtension;

            // Ensure the extension starts with a dot.
            if (!extension.StartsWith("."))
                extension = "." + extension;

            string fileName = $"OleObject_{i}{extension}";
            string outputPath = Path.Combine(outputFolder, fileName);

            // Save the OLE object data directly to a file.
            oleFormat.Save(outputPath);

            Console.WriteLine($"Extracted OLE object to: {outputPath}");
        }

        Console.WriteLine("Extraction complete.");
    }
}
