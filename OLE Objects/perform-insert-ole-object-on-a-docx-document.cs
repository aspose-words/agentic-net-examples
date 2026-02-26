using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleObjectExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph describing the OLE object that will be inserted.
        builder.Writeln("Embedded ZIP file as OLE object:");

        // Load the binary content of the file that we want to embed.
        // In this example we embed a ZIP archive, but any file can be used.
        byte[] zipBytes = File.ReadAllBytes(@"C:\Data\cat001.zip");

        // Use a memory stream to pass the file data to the InsertOleObject method.
        using (MemoryStream zipStream = new MemoryStream(zipBytes))
        {
            // Insert the OLE object as an icon.
            // Parameters:
            //   zipStream   – stream containing the file data.
            //   "Package"   – ProgId for a generic OLE package.
            //   true        – display the object as an icon.
            //   null        – no custom presentation image; Word will use a default icon.
            Shape oleShape = builder.InsertOleObject(zipStream, "Package", true, null);

            // Set the file name and display name that will appear when the user double‑clicks the icon.
            oleShape.OleFormat.OlePackage.FileName = "cat001.zip";
            oleShape.OleFormat.OlePackage.DisplayName = "Cat Archive.zip";
        }

        // Save the resulting document to a DOCX file.
        doc.Save(@"C:\Output\InsertOleObject.docx");
    }
}
