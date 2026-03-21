using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractOleObject
{
    static void Main()
    {
        // Use a document path relative to the executable directory.
        string baseDir = AppContext.BaseDirectory;
        string documentPath = Path.Combine(baseDir, "Sample.docx");
        string outputPath = Path.Combine(baseDir, "OleObject.bin");

        // If the sample document does not exist, create a minimal placeholder document.
        if (!File.Exists(documentPath))
        {
            Document placeholder = new Document();
            placeholder.Save(documentPath);
            Console.WriteLine($"Placeholder document created at: {documentPath}");
            Console.WriteLine("No OLE objects to extract. Exiting.");
            return;
        }

        // Ensure the output directory exists.
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

        // Load the document.
        Document doc = new Document(documentPath);

        // Retrieve the first shape that holds an OLE object.
        Shape oleShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
        if (oleShape == null || oleShape.OleFormat == null)
        {
            Console.WriteLine("No OLE object found in the document.");
            return;
        }

        // Save the OLE object's data to a binary file.
        using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
        {
            oleShape.OleFormat.Save(fileStream);
        }

        Console.WriteLine($"OLE object extracted successfully to: {outputPath}");
    }
}
