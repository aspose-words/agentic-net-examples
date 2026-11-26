using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare a sample source file that will be embedded as an OLE object.
        string sourceFilePath = Path.Combine(Path.GetTempPath(), "SampleSource.txt");
        File.WriteAllText(sourceFilePath, "This is the content of the source file.");

        // Create a new document and insert the OLE object.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape oleShape = builder.InsertOleObject(sourceFilePath, false, false, null);

        // Save the document containing the OLE object.
        string docPath = Path.Combine(Path.GetTempPath(), "OleDocument.docx");
        doc.Save(docPath);

        // Load the document back.
        Document loadedDoc = new Document(docPath);

        // Find the first shape that contains an OLE object.
        Shape foundOleShape = null;
        foreach (Shape shape in loadedDoc.GetChildNodes(NodeType.Shape, true))
        {
            // An OLE object is present if the shape has an OleFormat with a non‑null OlePackage.
            if (shape.OleFormat != null && shape.OleFormat.OlePackage != null)
            {
                foundOleShape = shape;
                break;
            }
        }

        if (foundOleShape != null)
        {
            // Access the OLE package and retrieve its original file name.
            OlePackage olePackage = foundOleShape.OleFormat.OlePackage;
            string oleFileName = olePackage.FileName; // Original source file name stored in the package.

            // Compare with the original source file name.
            string originalFileName = Path.GetFileName(sourceFilePath);
            bool namesMatch = string.Equals(oleFileName, originalFileName, StringComparison.OrdinalIgnoreCase);

            Console.WriteLine($"Original source file name: {originalFileName}");
            Console.WriteLine($"OLE package file name:    {oleFileName}");
            Console.WriteLine($"Names match: {namesMatch}");
        }
        else
        {
            Console.WriteLine("No OLE object found in the document.");
        }

        // Clean up temporary files.
        try { File.Delete(sourceFilePath); } catch { }
        try { File.Delete(docPath); } catch { }
    }
}
