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
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare some dummy data to embed as a generic OLE package.
        byte[] dummyData = System.Text.Encoding.UTF8.GetBytes("Hello from OLE package");
        using (MemoryStream dataStream = new MemoryStream(dummyData))
        {
            // Insert the OLE object into the document.
            // progId "Package" indicates a generic OLE package.
            // asIcon = true to display it as an icon; presentation stream is null.
            builder.InsertOleObject(dataStream, "Package", true, null);
        }

        // Optional: save the document to a temporary file (not required for extraction).
        string docTempPath = Path.Combine(Path.GetTempPath(), "OleDocument.docx");
        doc.Save(docTempPath);

        // Locate the first shape that contains an OLE object.
        Shape oleShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
        OleFormat oleFormat = oleShape?.OleFormat;

        if (oleFormat != null)
        {
            // Retrieve the raw binary data of the OLE object.
            byte[] rawData = oleFormat.GetRawData();

            // Write the raw data to a temporary file for external analysis.
            string rawDataTempPath = Path.Combine(Path.GetTempPath(), "OleObjectData.bin");
            File.WriteAllBytes(rawDataTempPath, rawData);

            // Output the location of the saved raw data.
            Console.WriteLine($"OLE raw data saved to: {rawDataTempPath}");
        }
        else
        {
            Console.WriteLine("No OLE object found in the document.");
        }
    }
}
