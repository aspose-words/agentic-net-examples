using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class ExtractOleObjects
{
    static void Main()
    {
        // Path to the source DOCX document that contains embedded OLE objects.
        string sourceDocPath = @"C:\Docs\SourceDocument.docx";

        // Folder where the extracted OLE files will be saved.
        string outputFolder = @"C:\Docs\ExtractedOleObjects";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the Word document.
        Document doc = new Document(sourceDocPath);

        // Retrieve all Shape nodes in the document (including OLE objects).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        foreach (Shape shape in shapes)
        {
            // Process only shapes that are OLE objects.
            if (shape.ShapeType == ShapeType.OleObject)
            {
                // Access the OleFormat of the shape.
                OleFormat oleFormat = shape.OleFormat;

                // Determine a suitable file name.
                // Use the suggested file name if available; otherwise generate one.
                string fileName = oleFormat.SuggestedFileName;
                if (string.IsNullOrEmpty(fileName))
                {
                    // Fallback to a generic name with the suggested extension.
                    fileName = "ExtractedObject" + oleFormat.SuggestedExtension;
                }

                // Combine the output folder with the file name.
                string outputPath = Path.Combine(outputFolder, fileName);

                // Save the embedded OLE object directly to a file.
                // This uses the OleFormat.Save(string) method provided by Aspose.Words.
                oleFormat.Save(outputPath);

                Console.WriteLine($"Extracted OLE object saved to: {outputPath}");
            }
        }
    }
}
