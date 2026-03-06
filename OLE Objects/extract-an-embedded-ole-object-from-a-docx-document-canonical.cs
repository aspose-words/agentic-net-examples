using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Ole;
using Aspose.Words.Tables;

class ExtractOleObjects
{
    static void Main()
    {
        // Path to the source DOCX file that contains embedded OLE objects.
        string sourceDocPath = @"C:\Docs\SourceDocument.docx";

        // Directory where extracted OLE files will be saved.
        string outputDir = @"C:\Docs\ExtractedOle";
        Directory.CreateDirectory(outputDir);

        // Load the Word document.
        Document doc = new Document(sourceDocPath);

        // Iterate through all Shape nodes in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int oleIndex = 0;

        foreach (Shape shape in shapes)
        {
            // Only process shapes that actually contain an OLE object.
            if (shape.ShapeType != ShapeType.OleObject)
                continue;

            // Access the OleFormat of the shape.
            OleFormat oleFormat = shape.OleFormat;

            // Skip linked OLE objects – they cannot be saved directly.
            if (oleFormat.IsLink)
                continue;

            // Determine a file name for the extracted object.
            string fileExtension = oleFormat.SuggestedExtension ?? ".bin";
            string extractedFilePath = Path.Combine(outputDir,
                $"OleObject_{oleIndex}{fileExtension}");

            // Save the embedded OLE data to the file using the provided Save method.
            oleFormat.Save(extractedFilePath);

            Console.WriteLine($"Extracted OLE object saved to: {extractedFilePath}");
            oleIndex++;
        }

        if (oleIndex == 0)
            Console.WriteLine("No embedded OLE objects were found in the document.");
    }
}
