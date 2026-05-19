using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and attach a DocumentBuilder to it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare some data to embed as an OLE object (a simple text file in this case).
        byte[] fileData = System.Text.Encoding.UTF8.GetBytes("Sample embedded content");
        using (MemoryStream dataStream = new MemoryStream(fileData))
        {
            // Insert the OLE object from the stream.
            // progId "Package" is used for generic package objects.
            // asIcon = true displays the OLE object as an icon.
            // presentation = null lets Aspose.Words choose a default icon.
            Shape oleShape = builder.InsertOleObject(dataStream, "Package", true, null);

            // Apply custom layout dimensions (width and height are in points).
            oleShape.Width = 150;   // 150 points wide
            oleShape.Height = 100;  // 100 points high
        }

        // Save the resulting document to the file system.
        doc.Save("OleObjectCustomSize.docx");
    }
}
