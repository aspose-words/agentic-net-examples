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

        // Create a minimal in‑memory ZIP file (just the ZIP header bytes).
        byte[] zipBytes = new byte[] { 0x50, 0x4B, 0x03, 0x04 };
        using (MemoryStream stream = new MemoryStream(zipBytes))
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
