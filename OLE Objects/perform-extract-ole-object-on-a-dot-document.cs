using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractOleFromDot
{
    static void Main()
    {
        // Path to the DOT (Word template) document.
        string dotPath = @"C:\Docs\Template.dot";

        // Folder where extracted OLE objects will be saved.
        string outputFolder = @"C:\Docs\ExtractedOle";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the DOT document.
        Document doc = new Document(dotPath);

        // Iterate through all shapes in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int oleIndex = 0;

        foreach (Shape shape in shapes)
        {
            // Only process shapes that contain an OLE object.
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue;

            // Skip linked OLE objects – they cannot be saved directly.
            if (oleFormat.IsLink)
                continue;

            // Determine a file name for the extracted object.
            string extension = oleFormat.SuggestedExtension ?? ".bin";
            string fileName = $"OleObject_{oleIndex}{extension}";
            string fullPath = Path.Combine(outputFolder, fileName);

            // Save the OLE object data to the file.
            oleFormat.Save(fullPath);

            Console.WriteLine($"Extracted OLE object to: {fullPath}");
            oleIndex++;
        }

        // Optionally, save the (unchanged) document if needed.
        // doc.Save(@"C:\Docs\Template_Processed.dot");
    }
}
