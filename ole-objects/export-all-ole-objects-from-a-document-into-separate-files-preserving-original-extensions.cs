using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Path to the source document. Adjust the file name as needed.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "Input.docx");
        if (!File.Exists(inputPath))
            return; // Exit if the document is not found.

        // Load the document.
        Document doc = new Document(inputPath);

        // Get all shapes in the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int oleIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Each shape that contains an OLE object has a non‑null OleFormat.
            OleFormat ole = shape.OleFormat;
            if (ole == null)
                continue;

            // Determine a file name for the extracted OLE object.
            // Use the suggested file name if available; otherwise create a generic one.
            string baseName = !string.IsNullOrEmpty(ole.SuggestedFileName)
                ? Path.GetFileNameWithoutExtension(ole.SuggestedFileName)
                : $"OleObject_{oleIndex}";

            // Preserve the original extension suggested by Aspose.Words.
            string extension = ole.SuggestedExtension ?? ".bin";

            string outputFileName = $"{baseName}{extension}";
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), outputFileName);

            // Save the OLE object directly to a file.
            ole.Save(outputPath);

            oleIndex++;
        }
    }
}
