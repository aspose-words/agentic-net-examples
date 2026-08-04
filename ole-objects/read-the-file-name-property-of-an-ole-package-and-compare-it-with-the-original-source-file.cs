using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary source file that will be embedded as an OLE package.
        string tempDir = Path.GetTempPath();
        string sourceFilePath = Path.Combine(tempDir, "SourceFile.txt");
        File.WriteAllText(sourceFilePath, "This is the content of the source file.");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Read the source file into a memory stream.
        byte[] sourceBytes = File.ReadAllBytes(sourceFilePath);
        using (MemoryStream sourceStream = new MemoryStream(sourceBytes))
        {
            // Insert the OLE package into the document.
            // progId "Package" indicates a generic OLE package.
            Shape oleShape = builder.InsertOleObject(sourceStream, "Package", false, null);

            // Set the OLE package's FileName property to the original file name.
            oleShape.OleFormat.OlePackage.FileName = Path.GetFileName(sourceFilePath);
        }

        // Save the document to a temporary location.
        string docPath = Path.Combine(tempDir, "OlePackageDemo.docx");
        doc.Save(docPath);

        // Load the document back.
        Document loadedDoc = new Document(docPath);

        // Retrieve the first shape that contains the OLE object.
        Shape loadedOleShape = (Shape)loadedDoc.GetChild(NodeType.Shape, 0, true);
        OlePackage loadedPackage = loadedOleShape.OleFormat.OlePackage;

        // Read the FileName property from the OLE package.
        string embeddedFileName = loadedPackage?.FileName ?? string.Empty;

        // Compare with the original source file name.
        string originalFileName = Path.GetFileName(sourceFilePath);
        bool namesMatch = string.Equals(embeddedFileName, originalFileName, StringComparison.OrdinalIgnoreCase);

        // Output the comparison result.
        Console.WriteLine($"Original file name: {originalFileName}");
        Console.WriteLine($"Embedded OLE package file name: {embeddedFileName}");
        Console.WriteLine($"Do the names match? {namesMatch}");
    }
}
