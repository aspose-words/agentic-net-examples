using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Load a file (e.g., a ZIP archive) that will be embedded as an OLE object.
        // Adjust the path to point to an existing file on your system.
        byte[] zipBytes = File.ReadAllBytes("cat001.zip");

        // Insert the OLE object from the memory stream.
        using (MemoryStream stream = new MemoryStream(zipBytes))
        {
            // Insert as an icon (asIcon = true) with no custom presentation image (presentation = null).
            Shape shape = builder.InsertOleObject(stream, "Package", true, null);

            // Optionally set the file name and display name that appear when the OLE object is opened.
            shape.OleFormat.OlePackage.FileName = "Package file name.zip";
            shape.OleFormat.OlePackage.DisplayName = "Package display name.zip";
        }

        // Save the resulting document to disk.
        doc.Save("OleObjectExample.docx");
    }
}
