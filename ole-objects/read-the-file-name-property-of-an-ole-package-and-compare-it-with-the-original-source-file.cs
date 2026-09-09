using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary folder for test files.
        string baseDir = Path.Combine(Environment.CurrentDirectory, "OleExample");
        Directory.CreateDirectory(baseDir);

        // Create a source file that will be embedded as an OLE package.
        string sourceFilePath = Path.Combine(baseDir, "sample.txt");
        File.WriteAllText(sourceFilePath, "This is a sample text file for OLE package testing.");

        // Load the source file into a byte array.
        byte[] sourceBytes = File.ReadAllBytes(sourceFilePath);

        // Create a new blank document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the OLE package into the document from the byte array.
        using (MemoryStream stream = new MemoryStream(sourceBytes))
        {
            // Insert as an OLE object of type "Package".
            Shape oleShape = builder.InsertOleObject(stream, "Package", true, null);

            // Set the OLE package's FileName property to the original file name.
            oleShape.OleFormat.OlePackage.FileName = Path.GetFileName(sourceFilePath);
        }

        // Save the document to disk.
        string docPath = Path.Combine(baseDir, "OleDocument.docx");
        doc.Save(docPath);

        // Load the document back from the file.
        Document loadedDoc = new Document(docPath);

        // Retrieve the first shape that contains an OLE object.
        Shape loadedOleShape = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                        .OfType<Shape>()
                                        .FirstOrDefault(s => s.OleFormat != null);

        if (loadedOleShape == null)
        {
            Console.WriteLine("No OLE object found in the loaded document.");
            return;
        }

        // Access the OLE package and read its FileName property.
        OlePackage olePackage = loadedOleShape.OleFormat.OlePackage;
        string oleFileName = olePackage?.FileName ?? string.Empty;

        // Compare the OLE package file name with the original source file name.
        string originalFileName = Path.GetFileName(sourceFilePath);
        bool namesMatch = string.Equals(oleFileName, originalFileName, StringComparison.Ordinal);

        // Output the comparison result.
        Console.WriteLine($"Original file name: {originalFileName}");
        Console.WriteLine($"OLE package file name: {oleFileName}");
        Console.WriteLine($"Names match: {namesMatch}");
    }
}
