using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExportOleObjects
{
    static void Main()
    {
        // Path to the source Word document (placed in the executable folder).
        string sourcePath = Path.Combine(AppContext.BaseDirectory, "InputDocument.docx");

        // Folder where extracted OLE files will be saved.
        string outputFolder = Path.Combine(AppContext.BaseDirectory, "ExtractedOleObjects");

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // If the source document does not exist, create a minimal one.
        if (!File.Exists(sourcePath))
        {
            var emptyDoc = new Document();
            emptyDoc.Save(sourcePath);
            Console.WriteLine($"Created placeholder document at: {sourcePath}");
        }

        // Load the Word document.
        Document doc = new Document(sourcePath);

        // Get all Shape nodes in the document (including those inside groups).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int oleIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Access the OLE data of the shape, if any.
            OleFormat ole = shape.OleFormat;
            if (ole == null)
                continue; // Not an OLE object.

            // Skip linked OLE objects; only extract embedded data.
            if (ole.IsLink)
                continue;

            // Determine a file name for the extracted object.
            // Prefer the suggested file name; otherwise build one using the suggested extension.
            string fileName = ole.SuggestedFileName;
            if (string.IsNullOrEmpty(fileName))
            {
                // Fallback: OleObject_0.bin, OleObject_1.bin, etc.
                string extension = ole.SuggestedExtension ?? ".bin";
                fileName = $"OleObject_{oleIndex}{extension}";
            }

            // Combine folder and file name.
            string outputPath = Path.Combine(outputFolder, fileName);

            // Save the OLE object data to the file system.
            ole.Save(outputPath);

            Console.WriteLine($"Extracted OLE object #{oleIndex} to: {outputPath}");
            oleIndex++;
        }

        Console.WriteLine("Extraction complete.");
    }
}
