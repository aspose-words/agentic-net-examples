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

        // Prepare some dummy data to embed as an OLE Package.
        byte[] dummyData = System.Text.Encoding.UTF8.GetBytes("Sample OLE binary content");
        using (MemoryStream dataStream = new MemoryStream(dummyData))
        {
            // Insert the OLE object into the document.
            // "Package" progId indicates a generic OLE package.
            // Insert as an icon (asIcon = true) with no custom presentation image.
            builder.InsertOleObject(dataStream, "Package", true, null);
        }

        // Retrieve the first shape that contains the OLE object.
        Shape oleShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
        OleFormat oleFormat = oleShape.OleFormat;

        // Get the raw binary data of the OLE object.
        byte[] rawData = oleFormat.GetRawData();

        // Save the raw data to a temporary file for external analysis.
        string tempFilePath = Path.Combine(Path.GetTempPath(), "OleObjectData.bin");
        File.WriteAllBytes(tempFilePath, rawData);

        // Output the location of the temporary file.
        Console.WriteLine($"OLE raw data saved to: {tempFilePath}");
    }
}
