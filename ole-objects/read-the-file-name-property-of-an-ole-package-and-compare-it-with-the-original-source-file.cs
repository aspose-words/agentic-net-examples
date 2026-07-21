using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class OlePackageFileNameExample
{
    public static void Main()
    {
        // Prepare dummy data to embed as an OLE package.
        byte[] dummyData = new byte[] { 0x50, 0x4B, 0x03, 0x04 }; // Minimal ZIP header bytes.
        using (MemoryStream dataStream = new MemoryStream(dummyData))
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert the OLE package into the document.
            // progId "Package" indicates a generic OLE package.
            Shape oleShape = builder.InsertOleObject(dataStream, "Package", true, null);

            // Set the OLE package file name property.
            string originalFileName = "dummy.zip";
            oleShape.OleFormat.OlePackage.FileName = originalFileName;

            // Save the document to disk.
            string docPath = Path.Combine(Directory.GetCurrentDirectory(), "OlePackage.docx");
            doc.Save(docPath);

            // Load the saved document.
            Document loadedDoc = new Document(docPath);

            // Retrieve the first shape that contains the OLE object.
            Shape loadedOleShape = (Shape)loadedDoc.GetChild(NodeType.Shape, 0, true);

            // Access the OLE package and read its FileName property.
            OlePackage loadedPackage = loadedOleShape.OleFormat.OlePackage;
            string loadedFileName = loadedPackage?.FileName ?? string.Empty;

            // Compare the stored file name with the original file name.
            bool namesMatch = string.Equals(loadedFileName, originalFileName, StringComparison.OrdinalIgnoreCase);

            // Output the comparison result.
            Console.WriteLine($"Original file name set: {originalFileName}");
            Console.WriteLine($"File name read from OLE package: {loadedFileName}");
            Console.WriteLine($"Do the names match? {namesMatch}");
        }
    }
}
