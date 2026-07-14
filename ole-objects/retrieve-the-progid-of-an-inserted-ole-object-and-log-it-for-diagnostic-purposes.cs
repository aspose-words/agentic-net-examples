using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare some dummy data to embed as an OLE object.
        byte[] dummyData = new byte[] { 0x01, 0x02, 0x03, 0x04 };
        using (MemoryStream stream = new MemoryStream(dummyData))
        {
            // Insert the OLE object using the "Package" ProgID.
            // Parameters: stream, progId, asIcon, presentation (null for default icon).
            Shape oleShape = builder.InsertOleObject(stream, "Package", true, null);

            // Retrieve the OleFormat from the inserted shape.
            OleFormat oleFormat = oleShape.OleFormat;

            // Get the ProgId of the OLE object.
            string progId = oleFormat.ProgId;

            // Log the ProgId for diagnostic purposes.
            Console.WriteLine($"Inserted OLE object's ProgId: {progId}");
        }

        // Save the document to verify the OLE object was inserted.
        doc.Save("InsertedOleObject.docx");
    }
}
