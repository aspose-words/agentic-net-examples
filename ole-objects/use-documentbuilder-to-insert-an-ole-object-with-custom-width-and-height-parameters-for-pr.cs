using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and attach a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare some dummy data to embed as an OLE package.
        byte[] dummyData = System.Text.Encoding.UTF8.GetBytes("Hello from OLE package");
        using (MemoryStream oleStream = new MemoryStream(dummyData))
        {
            // Insert the OLE object as an icon. The progId "Package" denotes a generic OLE package.
            Shape oleShape = builder.InsertOleObject(oleStream, "Package", true, null);

            // Apply custom layout dimensions (points). 1 point = 1/72 inch.
            oleShape.Width = 150;   // Width in points.
            oleShape.Height = 100;  // Height in points.
        }

        // Save the document to the file system.
        doc.Save("OleObjectCustomSize.docx");
    }
}
