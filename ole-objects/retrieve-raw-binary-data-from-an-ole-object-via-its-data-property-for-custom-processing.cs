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

        // Prepare some dummy data to embed as an OLE package.
        byte[] dummyData = System.Text.Encoding.UTF8.GetBytes("Hello, OLE object!");
        using (MemoryStream dataStream = new MemoryStream(dummyData))
        {
            // Insert the OLE object into the document.
            // progId "Package" indicates a generic OLE package.
            // asIcon = false to display the content directly.
            // presentation = null (default icon will be used if needed).
            builder.InsertOleObject(dataStream, "Package", false, null);
        }

        // Locate the first shape that contains the OLE object.
        Shape oleShape = doc.GetChildNodes(NodeType.Shape, true)
                            .OfType<Shape>()
                            .FirstOrDefault(s => s.OleFormat != null);

        if (oleShape != null && oleShape.OleFormat != null)
        {
            // Retrieve the raw binary data of the OLE object.
            byte[] rawData = oleShape.OleFormat.GetRawData();

            // Example custom processing: output the size of the raw data.
            Console.WriteLine($"OLE raw data length: {rawData.Length} bytes");

            // Optional: write the raw data to a file for verification.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "OleRawData.bin");
            File.WriteAllBytes(outputPath, rawData);
            Console.WriteLine($"Raw data saved to: {outputPath}");
        }

        // Save the document (optional, demonstrates usage of the Save method).
        string docPath = Path.Combine(Environment.CurrentDirectory, "OleDocument.docx");
        doc.Save(docPath);
        Console.WriteLine($"Document saved to: {docPath}");
    }
}
