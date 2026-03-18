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

        // Load a file to embed as an OLE object (replace with an existing file path).
        byte[] fileBytes = File.ReadAllBytes("cat001.zip");
        using (MemoryStream stream = new MemoryStream(fileBytes))
        {
            // Insert the OLE object into the document.
            // Parameters: data stream, ProgId ("Package" for generic OLE package), display as content (asIcon = false), no custom icon.
            Shape oleShape = builder.InsertOleObject(stream, "Package", false, null);

            // Retrieve the ProgId of the inserted OLE object.
            string progId = oleShape.OleFormat.ProgId;

            // Log the ProgId for diagnostic purposes.
            Console.WriteLine($"Inserted OLE object's ProgId: {progId}");
        }

        // Save the document to verify the insertion.
        doc.Save("OleProgId.docx");
    }
}
