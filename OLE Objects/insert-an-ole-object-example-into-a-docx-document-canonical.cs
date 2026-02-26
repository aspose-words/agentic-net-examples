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

        // Path to the file that will be embedded as an OLE object.
        // In this example we embed a ZIP archive, but any file can be used.
        string zipPath = "Data/cat001.zip";

        // Read the file into a memory stream.
        byte[] zipBytes = File.ReadAllBytes(zipPath);
        using (MemoryStream zipStream = new MemoryStream(zipBytes))
        {
            // Insert the OLE object.
            // Parameters:
            //   zipStream   – stream containing the file data.
            //   "Package"   – ProgID for a generic OLE package.
            //   true        – display the object as an icon.
            //   null        – no custom presentation image; a default icon will be used.
            Shape oleShape = builder.InsertOleObject(zipStream, "Package", true, null);

            // Set the file name and display name that Word will show for the OLE package.
            oleShape.OleFormat.OlePackage.FileName = "Package file name.zip";
            oleShape.OleFormat.OlePackage.DisplayName = "Package display name.zip";
        }

        // Save the resulting document.
        doc.Save("Output/InsertOleObject.docx");
    }
}
