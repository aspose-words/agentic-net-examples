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

        // Initialize a DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Path to the file that will be embedded as an OLE object (e.g., a ZIP archive).
        string zipPath = @"C:\Data\example.zip";

        // Load the file into a memory stream.
        byte[] zipBytes = File.ReadAllBytes(zipPath);
        using (MemoryStream zipStream = new MemoryStream(zipBytes))
        {
            // Insert the OLE object from the stream.
            // Parameters: stream, progId ("Package" for generic files), asIcon (true to show as icon), presentation (null uses default icon).
            Shape oleShape = builder.InsertOleObject(zipStream, "Package", true, null);

            // Set the OLE package's file name and display name.
            oleShape.OleFormat.OlePackage.FileName = "example.zip";
            oleShape.OleFormat.OlePackage.DisplayName = "Example ZIP Archive";
        }

        // Save the document as a DOCX file.
        doc.Save(@"C:\Output\OleObjectExample.docx");
    }
}
