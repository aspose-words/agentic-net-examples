using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractOleFromHtml
{
    static void Main()
    {
        // Path to the source HTML document.
        string htmlPath = @"C:\Docs\source.html";

        // Directory where extracted OLE objects will be saved.
        string outputDir = @"C:\Docs\ExtractedOleObjects";
        Directory.CreateDirectory(outputDir);

        // Load the HTML document.
        Document doc = new Document(htmlPath);

        int oleIndex = 0;

        // Iterate through all shapes in the document.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Check if the shape contains an embedded OLE object (not a link).
            if (shape.ShapeType == ShapeType.OleObject && shape.OleFormat != null && !shape.OleFormat.IsLink)
            {
                OleFormat ole = shape.OleFormat;

                // Determine a suitable file extension; fallback to .bin if unavailable.
                string extension = ole.SuggestedExtension ?? ".bin";

                // Build the output file name.
                string outputPath = Path.Combine(outputDir, $"OleObject_{oleIndex}{extension}");

                // Save the OLE object directly to a file.
                ole.Save(outputPath);

                oleIndex++;
            }
        }

        Console.WriteLine($"Extracted {oleIndex} OLE object(s) to \"{outputDir}\".");
    }
}
