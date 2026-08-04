using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class ExportOleObjects
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // Prepare a temporary folder for all generated files.
        // -----------------------------------------------------------------
        string baseDir = Directory.GetCurrentDirectory();
        string tempDir = Path.Combine(baseDir, "TempOle");
        Directory.CreateDirectory(tempDir);

        // -----------------------------------------------------------------
        // Create a simple text file that will be embedded as an OLE object.
        // -----------------------------------------------------------------
        string sampleTextPath = Path.Combine(tempDir, "sample.txt");
        File.WriteAllText(sampleTextPath, "This is a sample OLE embedded file.");

        // -----------------------------------------------------------------
        // Create a new Word document and embed the text file as an OLE object.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert the OLE object (embedded, not linked, not displayed as an icon).
        builder.InsertOleObject(sampleTextPath, false, false, null);

        // -----------------------------------------------------------------
        // Directory where extracted OLE objects will be saved.
        // -----------------------------------------------------------------
        string outputDir = Path.Combine(baseDir, "ExtractedOleObjects");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Iterate through all shapes in the document and export any OLE objects.
        // -----------------------------------------------------------------
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int oleIndex = 0;

        foreach (Shape shape in shapes)
        {
            OleFormat ole = shape.OleFormat;
            if (ole == null)
                continue; // Not an OLE object.

            // Skip linked OLE objects – they cannot be saved directly.
            if (ole.IsLink)
                continue;

            // Use the suggested extension (includes the leading dot) to build the file name.
            string extension = ole.SuggestedExtension ?? ".bin";
            string fileName = $"OleObject_{oleIndex}{extension}";
            string fullPath = Path.Combine(outputDir, fileName);

            // Save the OLE object to the file system.
            ole.Save(fullPath);
            oleIndex++;
        }

        // Optional: inform the user where the files were saved.
        Console.WriteLine($"Extracted {oleIndex} OLE object(s) to: {outputDir}");
    }
}
