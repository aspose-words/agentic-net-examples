using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare directories.
        string baseDir = Directory.GetCurrentDirectory();
        string dataDir = Path.Combine(baseDir, "Data");
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(dataDir);
        Directory.CreateDirectory(outputDir);

        // Create a simple source file that will be embedded as an OLE package.
        string sourceFilePath = Path.Combine(dataDir, "SampleFile.txt");
        File.WriteAllText(sourceFilePath, "This is a sample file for OLE package testing.");

        // Load the source file bytes.
        byte[] sourceBytes = File.ReadAllBytes(sourceFilePath);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the OLE package into the document from the byte array.
        using (MemoryStream stream = new MemoryStream(sourceBytes))
        {
            // Insert as an OLE package (ProgId "Package") and display it as an icon.
            Shape oleShape = builder.InsertOleObject(stream, "Package", true, null);

            // Set the OLE package file name to the original source file name.
            oleShape.OleFormat.OlePackage.FileName = Path.GetFileName(sourceFilePath);
        }

        // Save the document containing the OLE object.
        string docPath = Path.Combine(outputDir, "OlePackageDocument.docx");
        doc.Save(docPath);

        // Load the saved document.
        Document loadedDoc = new Document(docPath);

        // Retrieve the first shape that contains the OLE object.
        Shape loadedOleShape = (Shape)loadedDoc.GetChild(NodeType.Shape, 0, true);

        // Access the OLE package and read its FileName property.
        string oleFileName = loadedOleShape.OleFormat.OlePackage.FileName;

        // Compare the OLE package file name with the original source file name.
        string originalFileName = Path.GetFileName(sourceFilePath);
        bool namesMatch = string.Equals(oleFileName, originalFileName, StringComparison.OrdinalIgnoreCase);

        // Output the comparison result.
        Console.WriteLine($"Original source file name: {originalFileName}");
        Console.WriteLine($"OLE package file name:    {oleFileName}");
        Console.WriteLine($"Names match: {namesMatch}");
    }
}
