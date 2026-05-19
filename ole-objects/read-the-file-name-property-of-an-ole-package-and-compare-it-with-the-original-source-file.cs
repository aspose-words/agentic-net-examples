using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class OlePackageFileNameExample
{
    public static void Main()
    {
        // Prepare a temporary folder for input and output files.
        string workDir = Path.Combine(Path.GetTempPath(), "OlePackageExample");
        Directory.CreateDirectory(workDir);

        // Create a simple source file that will be embedded as an OLE package.
        string sourceFilePath = Path.Combine(workDir, "SampleData.txt");
        File.WriteAllText(sourceFilePath, "This is sample content for the OLE package.");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Read the source file into a byte array and insert it as an OLE package.
        byte[] sourceBytes = File.ReadAllBytes(sourceFilePath);
        using (MemoryStream stream = new MemoryStream(sourceBytes))
        {
            // Insert the OLE object. "Package" indicates an OLE package type.
            Shape oleShape = builder.InsertOleObject(stream, "Package", true, null);

            // Set the OLE package's FileName property to the original file name.
            oleShape.OleFormat.OlePackage.FileName = Path.GetFileName(sourceFilePath);
        }

        // Save the document to disk.
        string docPath = Path.Combine(workDir, "OlePackageDocument.docx");
        doc.Save(docPath);

        // Load the saved document.
        Document loadedDoc = new Document(docPath);

        // Retrieve the first shape that contains an OLE object.
        Shape loadedOleShape = (Shape)loadedDoc.GetChild(NodeType.Shape, 0, true);
        OlePackage olePackage = loadedOleShape.OleFormat.OlePackage;

        // Read the FileName property from the OLE package.
        string oleFileName = olePackage?.FileName ?? string.Empty;

        // Compare the OLE package file name with the original source file name.
        bool namesMatch = string.Equals(oleFileName, Path.GetFileName(sourceFilePath), StringComparison.Ordinal);

        // Output the result.
        Console.WriteLine($"Original source file name: {Path.GetFileName(sourceFilePath)}");
        Console.WriteLine($"OLE package file name:    {oleFileName}");
        Console.WriteLine($"Names match: {namesMatch}");
    }
}
