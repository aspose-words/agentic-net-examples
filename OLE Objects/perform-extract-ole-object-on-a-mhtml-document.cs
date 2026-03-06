using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractOleFromMhtml
{
    static void Main()
    {
        // Path to the source MHTML document.
        string sourceMhtmlPath = @"C:\Docs\source.mht";

        // Folder where extracted OLE objects will be saved.
        string outputFolder = @"C:\Docs\ExtractedOle";
        Directory.CreateDirectory(outputFolder);

        // Load the MHTML document.
        Document doc = new Document(sourceMhtmlPath);

        // Iterate through all shapes in the document.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Check if the shape contains an OLE object.
            OleFormat ole = shape.OleFormat;
            if (ole == null)
                continue; // Not an OLE object, skip.

            // Determine a file name for the extracted object.
            // Use the suggested extension if available; otherwise default to .bin.
            string extension = string.IsNullOrEmpty(ole.SuggestedExtension) ? ".bin" : ole.SuggestedExtension;
            string fileName = $"OleObject_{Guid.NewGuid()}{extension}";
            string outputPath = Path.Combine(outputFolder, fileName);

            // Save the OLE object data to the file.
            // This will throw if the OLE object is a linked object; we only handle embedded objects.
            ole.Save(outputPath);

            Console.WriteLine($"Extracted OLE object saved to: {outputPath}");
        }
    }
}
