using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class OlePackageEnumerator
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare a simple text file content to embed as an OLE package.
        byte[] fileBytes = System.Text.Encoding.UTF8.GetBytes("Hello from embedded OLE package!");

        // Insert the OLE package into the document.
        using (MemoryStream stream = new MemoryStream(fileBytes))
        {
            Shape shape = builder.InsertOleObject(stream, "Package", true, null);
            // Set package properties for identification.
            shape.OleFormat.OlePackage.FileName = "Sample.txt";
            shape.OleFormat.OlePackage.DisplayName = "Sample OLE Package";
        }

        // Save the document to a temporary file (required for Aspose.Words to process shapes).
        string docPath = Path.Combine(Path.GetTempPath(), "OlePackageDemo.docx");
        doc.Save(docPath);

        // Reload the document (demonstrates loading with default options).
        Document loadedDoc = new Document(docPath);

        // Iterate over all shapes and enumerate OLE packages.
        var oleShapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                 .Cast<Shape>()
                                 .Where(s => s.OleFormat != null && s.OleFormat.OlePackage != null);

        foreach (var shape in oleShapes)
        {
            OlePackage package = shape.OleFormat.OlePackage;
            Console.WriteLine("Found OLE Package:");
            Console.WriteLine($"  FileName   : {package.FileName}");
            Console.WriteLine($"  DisplayName: {package.DisplayName}");
        }

        // Clean up temporary file.
        if (File.Exists(docPath))
        {
            File.Delete(docPath);
        }
    }
}
