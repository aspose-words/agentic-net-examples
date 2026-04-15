using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare OLE object data (a simple text file) in memory.
        byte[] oleData = System.Text.Encoding.UTF8.GetBytes("Sample OLE package content");
        using (MemoryStream oleStream = new MemoryStream(oleData))
        {
            // Insert the OLE object using the generic "Package" progId.
            // asIcon = true displays the object as an icon; presentation stream is null.
            Shape oleShape = builder.InsertOleObject(oleStream, "Package", true, null);

            // Apply custom layout dimensions (points).
            oleShape.Width = 200;   // Width in points.
            oleShape.Height = 100;  // Height in points.
        }

        // Save the document to the file system.
        doc.Save("OleObjectWithSize.docx");
    }
}
