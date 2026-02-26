using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;

class ExtractOleFromRtf
{
    static void Main()
    {
        // Path to the source RTF document.
        string rtfPath = @"Input.rtf";

        // Directory where extracted OLE objects will be saved.
        string outputFolder = @"ExtractedOle";
        Directory.CreateDirectory(outputFolder);

        // Load the RTF document. No special options are required for OLE extraction.
        RtfLoadOptions loadOptions = new RtfLoadOptions();
        Document doc = new Document(rtfPath, loadOptions);

        // Iterate through all Shape nodes in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int oleCounter = 0;

        foreach (Shape shape in shapes)
        {
            // Only shapes that contain an OLE object have a non‑null OleFormat.
            OleFormat ole = shape.OleFormat;
            if (ole == null)
                continue;

            // Determine a suitable file extension for the embedded object.
            // If the document cannot suggest one, fall back to a generic binary extension.
            string extension = ole.SuggestedExtension ?? ".bin";

            // Build a unique file name for each extracted object.
            string outputPath = Path.Combine(outputFolder,
                                            $"OleObject_{oleCounter}{extension}");

            // Save the OLE object data directly to the file.
            ole.Save(outputPath);

            Console.WriteLine($"Extracted OLE object to: {outputPath}");
            oleCounter++;
        }

        Console.WriteLine("Extraction complete.");
    }
}
