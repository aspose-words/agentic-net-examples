using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary folder for the example files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // Create a dummy source file that will be embedded as an OLE package.
        string sourceFilePath = Path.Combine(workDir, "sample.zip");
        File.WriteAllBytes(sourceFilePath, new byte[] { 0x50, 0x4B, 0x03, 0x04 }); // minimal ZIP header

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the source file as an OLE package.
        using (FileStream fileStream = File.OpenRead(sourceFilePath))
        {
            // Insert as an OLE object of type "Package". Display it as an icon.
            Shape oleShape = builder.InsertOleObject(fileStream, "Package", true, null);

            // Set the OLE package's FileName property to the original file name.
            oleShape.OleFormat.OlePackage.FileName = Path.GetFileName(sourceFilePath);
        }

        // Save the document to disk.
        string docPath = Path.Combine(workDir, "OlePackageExample.docx");
        doc.Save(docPath);

        // Load the document back.
        Document loadedDoc = new Document(docPath);

        // Retrieve the first shape that contains the OLE object.
        Shape loadedShape = (Shape)loadedDoc.GetChild(NodeType.Shape, 0, true);

        // Access the OLE package and read its FileName property.
        string embeddedFileName = loadedShape.OleFormat?.OlePackage?.FileName;

        // Compare the embedded file name with the original source file name.
        string originalFileName = Path.GetFileName(sourceFilePath);
        bool namesMatch = string.Equals(embeddedFileName, originalFileName, StringComparison.OrdinalIgnoreCase);

        // Output the result.
        Console.WriteLine($"Original file name: {originalFileName}");
        Console.WriteLine($"Embedded OLE package file name: {embeddedFileName}");
        Console.WriteLine($"Names match: {namesMatch}");
    }
}
