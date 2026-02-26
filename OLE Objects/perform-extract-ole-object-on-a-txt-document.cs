using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;

class ExtractOleFromTxt
{
    static void Main()
    {
        // Path to the input TXT document that may contain embedded OLE objects.
        string inputPath = @"C:\Docs\sample.txt";

        // Load the TXT document using TxtLoadOptions to preserve OLE objects.
        TxtLoadOptions loadOptions = new TxtLoadOptions();
        Document doc = new Document(inputPath, loadOptions);

        // Iterate through all Shape nodes in the document. OLE objects are stored inside Shape nodes.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Check if the shape actually contains an OLE object.
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue; // No OLE data in this shape.

            // Determine a suitable file name for the extracted object.
            // Use the suggested extension provided by Aspose.Words.
            string outputFileName = $"ExtractedObject_{Guid.NewGuid()}{oleFormat.SuggestedExtension}";
            string outputPath = Path.Combine(@"C:\ExtractedOle", outputFileName);

            // Ensure the output directory exists.
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

            // Save the OLE object directly to a file.
            // This uses the OleFormat.Save(string) method.
            oleFormat.Save(outputPath);

            Console.WriteLine($"Extracted OLE object saved to: {outputPath}");
        }
    }
}
