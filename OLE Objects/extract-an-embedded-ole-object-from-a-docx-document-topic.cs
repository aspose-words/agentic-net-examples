using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class OleExtractor
{
    // Extracts all embedded OLE objects from a DOCX file and saves them to the specified folder.
    public static void ExtractOleObjects(string docxPath, string outputFolder)
    {
        // Load the Word document from the file system.
        Document doc = new Document(docxPath);

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Get all Shape nodes in the document (including those inside headers/footers).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        int oleIndex = 0;
        foreach (Shape shape in shapes)
        {
            // Process only shapes that contain an OLE object.
            if (shape.ShapeType == ShapeType.OleObject && shape.OleFormat != null)
            {
                OleFormat oleFormat = shape.OleFormat;

                // Determine a file name for the extracted object.
                // Use the suggested file name if available; otherwise, create a generic name.
                string fileName = !string.IsNullOrEmpty(oleFormat.SuggestedFileName)
                    ? oleFormat.SuggestedFileName
                    : $"OleObject_{oleIndex}{oleFormat.SuggestedExtension}";

                // Combine the output folder with the file name.
                string outputPath = Path.Combine(outputFolder, fileName);

                // Save the OLE object directly to a file.
                // This uses the OleFormat.Save(string) method provided by Aspose.Words.
                oleFormat.Save(outputPath);

                Console.WriteLine($"Extracted OLE object to: {outputPath}");
                oleIndex++;
            }
        }
    }

    // Example usage.
    public static void Main()
    {
        // Path to the source DOCX document containing embedded OLE objects.
        string sourceDocx = @"C:\Docs\SampleWithOle.docx";

        // Folder where extracted OLE files will be saved.
        string destinationFolder = @"C:\ExtractedOle";

        ExtractOleObjects(sourceDocx, destinationFolder);
    }
}
