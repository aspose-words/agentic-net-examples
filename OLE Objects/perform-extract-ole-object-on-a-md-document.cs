using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class OleExtractor
{
    // Extracts all embedded OLE objects from a Word document and saves them to the specified folder.
    public static void ExtractOleObjects(string documentPath, string outputFolder)
    {
        // Load the Word document.
        Document doc = new Document(documentPath);

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Iterate through all Shape nodes in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int oleIndex = 0;

        foreach (Shape shape in shapes)
        {
            // Only process shapes that contain OLE data.
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue;

            // Determine a file name for the extracted OLE object.
            // Use the suggested extension if available; otherwise default to .bin.
            string extension = string.IsNullOrEmpty(oleFormat.SuggestedExtension)
                ? ".bin"
                : oleFormat.SuggestedExtension;

            string fileName = Path.Combine(outputFolder,
                $"OleObject_{oleIndex}{extension}");

            // Save the OLE object directly to a file.
            // This works for embedded objects; linked objects will throw InvalidOperationException.
            try
            {
                oleFormat.Save(fileName);
                Console.WriteLine($"Extracted OLE object to: {fileName}");
            }
            catch (InvalidOperationException)
            {
                // Linked OLE objects cannot be saved; skip them or handle as needed.
                Console.WriteLine($"Skipped linked OLE object at index {oleIndex}.");
            }

            oleIndex++;
        }
    }

    // Example usage.
    static void Main()
    {
        string docPath = @"C:\Docs\SampleWithOle.docx";
        string outputDir = @"C:\Docs\ExtractedOle";

        ExtractOleObjects(docPath, outputDir);
    }
}
