using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Path to the file that will be embedded as an OLE package.
        // Replace with the actual file you want to embed.
        string sourceFilePath = @"C:\Data\cat001.zip";

        // Read the file into a byte array.
        byte[] fileBytes = File.ReadAllBytes(sourceFilePath);

        // Insert the OLE object from a memory stream.
        using (MemoryStream stream = new MemoryStream(fileBytes))
        {
            // "Package" progId creates a generic OLE package.
            // The third argument (true) inserts the object as an icon.
            Shape oleShape = builder.InsertOleObject(stream, "Package", true, null);

            // Preserve the original file name and extension in the OLE package metadata.
            string fileName = Path.GetFileName(sourceFilePath);
            oleShape.OleFormat.OlePackage.FileName = fileName;
            oleShape.OleFormat.OlePackage.DisplayName = fileName;
        }

        // Save the document to disk.
        doc.Save(@"C:\Output\OlePackage.docx");
    }
}
