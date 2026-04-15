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

        // Prepare some dummy data to embed as an OLE package.
        byte[] dummyData = System.Text.Encoding.UTF8.GetBytes("Sample OLE content");
        using (MemoryStream dataStream = new MemoryStream(dummyData))
        {
            // Insert the OLE object using the provided InsertOleObject method.
            // progId "Package" denotes a generic OLE package.
            Shape oleShape = builder.InsertOleObject(dataStream, "Package", false, null);

            // Retrieve the ProgId of the inserted OLE object.
            string progId = oleShape.OleFormat.ProgId;

            // Log the ProgId for diagnostic purposes.
            Console.WriteLine($"Inserted OLE object's ProgId: {progId}");
        }

        // Save the document to verify that the OLE object was added.
        doc.Save("OleProgIdOutput.docx");
    }
}
