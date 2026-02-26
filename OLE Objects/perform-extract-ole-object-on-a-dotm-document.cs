using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class OleExtractor
{
    // Extracts all embedded OLE objects from a DOTM file and saves them to the specified folder.
    static void Main()
    {
        // Path to the source DOTM document.
        string sourcePath = @"C:\Docs\Template.dotm";

        // Folder where extracted OLE files will be saved.
        string outputFolder = @"C:\Docs\ExtractedOle";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the DOTM document.
        Document doc = new Document(sourcePath);

        // Iterate through all Shape nodes in the document.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Only shapes that contain OLE data have a non‑null OleFormat.
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue;

            // Skip linked OLE objects – they cannot be saved directly.
            if (oleFormat.IsLink)
                continue;

            // Determine a file name for the extracted object.
            // Use the suggested file name and extension if available.
            string suggestedName = oleFormat.SuggestedFileName;
            if (string.IsNullOrEmpty(suggestedName))
                suggestedName = "OleObject";

            // Ensure the file name has the proper extension.
            string extension = oleFormat.SuggestedExtension;
            if (!string.IsNullOrEmpty(extension) && !suggestedName.EndsWith(extension, StringComparison.OrdinalIgnoreCase))
                suggestedName += extension;

            // Build the full output path.
            string outputPath = Path.Combine(outputFolder, suggestedName);

            // Save the OLE object data to the file.
            oleFormat.Save(outputPath);
        }

        Console.WriteLine("OLE extraction completed.");
    }
}
