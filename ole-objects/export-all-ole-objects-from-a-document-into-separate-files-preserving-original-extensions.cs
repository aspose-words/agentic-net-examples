using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Path to the source Word document containing OLE objects.
        // Ensure that the file "Input.docx" exists in the current working directory.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "Input.docx");

        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            Console.WriteLine("Please place a Word document named 'Input.docx' in the current directory and retry.");
            return;
        }

        // Directory where extracted OLE files will be saved.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "ExtractedOle");
        Directory.CreateDirectory(outputDir);

        // Load the document.
        Document doc = new Document(inputPath);

        // Get all shape nodes in the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int oleIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Access the OLE format of the shape.
            OleFormat ole = shape.OleFormat;
            if (ole == null)
                continue; // Not an OLE object.

            // Skip linked OLE objects; only extract embedded ones.
            if (ole.IsLink)
                continue;

            // Determine a file name using the suggested extension (or fallback to .bin).
            string extension = ole.SuggestedExtension ?? ".bin";
            string fileName = $"OleObject_{oleIndex}{extension}";
            string fullPath = Path.Combine(outputDir, fileName);

            // Save the OLE object data to the file.
            ole.Save(fullPath);
            Console.WriteLine($"Extracted OLE object to: {fullPath}");

            oleIndex++;
        }

        if (oleIndex == 0)
        {
            Console.WriteLine("No embedded OLE objects were found in the document.");
        }
    }
}
