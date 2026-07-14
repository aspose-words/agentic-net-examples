using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class OleObjectInsertionDemo
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare a simple byte array to act as OLE data.
        byte[] dummyData = new byte[] { 0x00, 0x01, 0x02, 0x03 };
        using (MemoryStream oleStream = new MemoryStream(dummyData))
        {
            // Insert the OLE object into the document.
            // Parameters: stream, progId ("Package" for generic OLE package), asIcon = true, presentation = null.
            Shape oleShape = builder.InsertOleObject(oleStream, "Package", true, null);

            // Verify that the insertion returned a non‑null shape and that it contains an OleFormat.
            bool insertionSuccessful = oleShape != null && oleShape.OleFormat != null;

            // Output the verification result.
            Console.WriteLine("OLE object insertion successful: " + insertionSuccessful);
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OleObjectDemo.docx");
        doc.Save(outputPath);
    }
}
