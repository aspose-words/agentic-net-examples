using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class OleExtractor
{
    public static void Main()
    {
        // Path to the source Word document containing OLE objects.
        const string sourceDocPath = "InputDocument.docx";

        // Folder where extracted OLE objects will be saved.
        const string outputFolder = "ExtractedOleObjects";

        // Ensure the source document exists.
        if (!File.Exists(sourceDocPath))
        {
            Console.WriteLine($"Source document not found: {sourceDocPath}");
            return;
        }

        // Load the Word document using Aspose.Words.
        Document doc = new Document(sourceDocPath);

        // Create the output folder if it does not exist.
        Directory.CreateDirectory(outputFolder);

        // Iterate over all shapes in the document.
        var shapes = doc.GetChildNodes(NodeType.Shape, true);
        int oleIndex = 0;

        foreach (Shape shape in shapes)
        {
            // Access the OLE format of the shape.
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue; // Not an OLE object.

            // Retrieve raw OLE data.
            byte[] oleData = oleFormat.GetRawData();

            // Determine a file name for the extracted object.
            string fileName = oleFormat.SuggestedFileName;
            if (string.IsNullOrEmpty(fileName))
            {
                // Fallback to a generated name using the suggested extension.
                string extension = oleFormat.SuggestedExtension ?? ".bin";
                fileName = $"OleObject_{oleIndex}{extension}";
            }

            // Build the full path for the output file.
            string outputPath = Path.Combine(outputFolder, fileName);

            // Write the OLE data to the file.
            File.WriteAllBytes(outputPath, oleData);

            Console.WriteLine($"Extracted OLE object saved to: {outputPath}");

            oleIndex++;
        }

        // Indicate completion.
        Console.WriteLine("OLE extraction completed.");
    }
}
