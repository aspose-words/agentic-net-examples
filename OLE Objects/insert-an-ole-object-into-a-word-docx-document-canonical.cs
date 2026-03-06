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

        // Initialize a DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph describing the OLE object.
        builder.Writeln("Embedded ZIP file as OLE object:");

        // Load the ZIP file bytes into a memory stream.
        // Replace the path with the actual location of your ZIP file.
        byte[] zipBytes = File.ReadAllBytes("cat001.zip");
        using (MemoryStream zipStream = new MemoryStream(zipBytes))
        {
            // Insert the OLE object as an icon.
            // progId "Package" denotes a generic OLE package.
            // asIcon = true displays the object as an icon.
            // presentation = null uses the default icon.
            Shape oleShape = builder.InsertOleObject(zipStream, "Package", true, null);

            // Set the file name and display name for the embedded package.
            oleShape.OleFormat.OlePackage.FileName = "cat001.zip";
            oleShape.OleFormat.OlePackage.DisplayName = "Sample ZIP archive";
        }

        // Save the document to a DOCX file.
        doc.Save("OleObjectExample.docx");
    }
}
