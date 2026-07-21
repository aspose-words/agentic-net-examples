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
        byte[] dummyData = System.Text.Encoding.UTF8.GetBytes("Hello, OLE!");
        using (MemoryStream stream = new MemoryStream(dummyData))
        {
            // Insert the OLE object. Use "Package" as the ProgId and display it as an icon.
            Shape oleShape = builder.InsertOleObject(stream, "Package", true, null);

            // Retrieve the ProgId of the inserted OLE object.
            string progId = oleShape.OleFormat.ProgId;

            // Log the ProgId.
            Console.WriteLine($"Inserted OLE object's ProgId: {progId}");
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OleProgIdExample.docx");
        doc.Save(outputPath);
    }
}
