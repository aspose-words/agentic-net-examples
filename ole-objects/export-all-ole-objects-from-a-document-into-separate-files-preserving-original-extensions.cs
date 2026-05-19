using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class ExportOleObjects
{
    public static void Main()
    {
        // Path to the input Word document.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "Input.docx");

        // Directory where extracted OLE files will be saved.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "ExtractedOle");
        Directory.CreateDirectory(outputDir);

        // Load the document.
        Document doc = new Document(inputPath);

        // Get all shapes in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        int oleIndex = 0;
        foreach (Shape shape in shapes)
        {
            OleFormat ole = shape.OleFormat;
            if (ole == null)
                continue;

            // Skip linked OLE objects because Save throws for them.
            if (ole.IsLink)
                continue;

            // Determine file extension; fallback to .bin if not provided.
            string extension = ole.SuggestedExtension;
            if (string.IsNullOrEmpty(extension))
                extension = ".bin";

            // Build output file name.
            string fileName = $"OleObject_{oleIndex}{extension}";
            string outputPath = Path.Combine(outputDir, fileName);

            // Save the OLE object to the file.
            ole.Save(outputPath);
            oleIndex++;
        }
    }
}
