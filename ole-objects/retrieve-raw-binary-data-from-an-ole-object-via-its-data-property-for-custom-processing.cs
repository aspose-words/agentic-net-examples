using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class RetrieveOleRawData
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare some sample data to embed as an OLE object.
        byte[] sampleData = System.Text.Encoding.UTF8.GetBytes("Sample OLE content");
        using (MemoryStream oleStream = new MemoryStream(sampleData))
        {
            // Insert the OLE object into the document as a package.
            // Parameters: stream, progId ("Package"), asIcon = false, presentation = null.
            builder.InsertOleObject(oleStream, "Package", false, null);
        }

        // Save the document (optional, just to have a file on disk).
        string docPath = "OleDocument.docx";
        doc.Save(docPath);

        // Load the document back (demonstrates the load step).
        Document loadedDoc = new Document(docPath);

        // Iterate through all shapes to find OLE objects.
        foreach (Shape shape in loadedDoc.GetChildNodes(NodeType.Shape, true).OfType<Shape>())
        {
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat != null)
            {
                // Retrieve the raw binary data of the OLE object.
                byte[] rawData = oleFormat.GetRawData();

                // Example custom processing: write the raw data to a file.
                string rawDataPath = $"OleRawData_{Guid.NewGuid():N}.bin";
                File.WriteAllBytes(rawDataPath, rawData);

                // Output information to the console.
                Console.WriteLine($"OLE object ProgId: {oleFormat.ProgId}");
                Console.WriteLine($"Raw data length: {rawData.Length} bytes");
                Console.WriteLine($"Raw data saved to: {rawDataPath}");
            }
        }
    }
}
