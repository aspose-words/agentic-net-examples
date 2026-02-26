using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleObjectExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Path to the file that will be embedded as an OLE object.
        // In this example we embed a ZIP archive, but any file type can be used.
        string zipPath = @"C:\Data\cat001.zip";

        // Read the file into a byte array.
        byte[] zipBytes = File.ReadAllBytes(zipPath);

        // Insert the OLE object from a memory stream.
        // Parameters:
        //   stream   – the data stream containing the file.
        //   progId   – "Package" tells Aspose.Words to treat the data as a generic OLE package.
        //   asIcon   – true to display the object as an icon.
        //   presentation – null to use the default icon.
        using (MemoryStream zipStream = new MemoryStream(zipBytes))
        {
            Shape oleShape = builder.InsertOleObject(zipStream, "Package", true, null);

            // Set the file name and display name that Word will show when the user opens the OLE object.
            oleShape.OleFormat.OlePackage.FileName = "Package file name.zip";
            oleShape.OleFormat.OlePackage.DisplayName = "Package display name.zip";
        }

        // Save the document to disk.
        doc.Save(@"C:\Output\InsertOleObject.docx");
    }
}
