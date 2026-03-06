using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractOleFromMarkdown
{
    static void Main()
    {
        // Path to the Markdown file to process
        string markdownPath = @"C:\Docs\sample.md";

        // Directory where extracted OLE objects will be saved
        string outputFolder = @"C:\Docs\ExtractedOle";
        Directory.CreateDirectory(outputFolder);

        // Load the Markdown document (Aspose.Words can load .md directly)
        Document doc = new Document(markdownPath);

        // Find all shapes that may contain OLE objects
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int oleCounter = 0;

        foreach (Shape shape in shapeNodes)
        {
            OleFormat ole = shape.OleFormat;
            if (ole == null)
                continue; // Not an OLE object

            // Skip linked OLE objects – Save throws InvalidOperationException for them
            if (ole.IsLink)
                continue;

            // Build a file name using the suggested extension for the embedded object
            string oleFilePath = Path.Combine(
                outputFolder,
                $"OleObject_{oleCounter}{ole.SuggestedExtension}");

            // Save the embedded OLE data directly to a file
            ole.Save(oleFilePath);
            oleCounter++;
        }
    }
}
