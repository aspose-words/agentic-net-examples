using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class RetrieveOleRawData
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare some sample binary data to embed as an OLE package.
        byte[] sampleData = System.Text.Encoding.UTF8.GetBytes("Hello, this is sample OLE data.");
        using (MemoryStream dataStream = new MemoryStream(sampleData))
        {
            // Insert the binary data as an OLE object (Package) displayed as an icon.
            // Parameters: stream, progId ("Package"), asIcon = true, presentation = null.
            builder.InsertOleObject(dataStream, "Package", true, null);
        }

        // Retrieve the first shape which should contain the OLE object we just inserted.
        Shape oleShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
        if (oleShape?.OleFormat != null)
        {
            OleFormat oleFormat = oleShape.OleFormat;

            // Get the raw binary data of the OLE object.
            byte[] rawData = oleFormat.GetRawData();

            // Example custom processing: write the raw data length and its first few bytes to console.
            Console.WriteLine($"OLE raw data length: {rawData.Length}");
            int previewLength = Math.Min(20, rawData.Length);
            string preview = BitConverter.ToString(rawData, 0, previewLength);
            Console.WriteLine($"First {previewLength} bytes: {preview}");
        }
        else
        {
            Console.WriteLine("No OLE object found in the document.");
        }

        // Optionally save the document to verify the OLE object persists.
        doc.Save("OleObjectDocument.docx");
    }
}
