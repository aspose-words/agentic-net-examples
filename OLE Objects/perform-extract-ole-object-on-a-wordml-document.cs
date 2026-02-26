using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractOleObjects
{
    static void Main()
    {
        // Path to the WORDML (WordprocessingML) document.
        string inputPath = @"C:\Docs\SampleWordML.xml";

        // Folder where extracted OLE files will be saved.
        string outputFolder = @"C:\Docs\ExtractedOleObjects";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the WORDML document.
        Document doc = new Document(inputPath);

        // Iterate through all shapes in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int oleIndex = 0;

        foreach (Shape shape in shapes)
        {
            // Only process shapes that contain an OLE object.
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue;

            // Determine a file name for the extracted OLE object.
            // Use the suggested extension if available; otherwise default to .bin.
            string extension = string.IsNullOrEmpty(oleFormat.SuggestedExtension)
                ? ".bin"
                : oleFormat.SuggestedExtension;

            string fileName = $"OleObject_{oleIndex}{extension}";
            string outputPath = Path.Combine(outputFolder, fileName);

            // Save the OLE object data to the file.
            oleFormat.Save(outputPath);

            Console.WriteLine($"Extracted OLE object #{oleIndex} to: {outputPath}");
            oleIndex++;
        }

        Console.WriteLine("Extraction complete.");
    }
}
