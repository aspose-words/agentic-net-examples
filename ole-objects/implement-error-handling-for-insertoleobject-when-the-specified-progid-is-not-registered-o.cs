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

        // Prepare some dummy data to embed as an OLE object.
        byte[] dummyData = new byte[] { 0x00, 0x01, 0x02, 0x03 };
        using (MemoryStream dataStream = new MemoryStream(dummyData))
        {
            // Attempt to insert an OLE object with a ProgId that is unlikely to be registered.
            string invalidProgId = "Invalid.ProgId";

            try
            {
                // InsertOleObject may throw an exception if the ProgId is not found.
                builder.InsertOleObject(dataStream, invalidProgId, asIcon: false, presentation: null);
                Console.WriteLine("OLE object inserted successfully.");
            }
            catch (Exception ex)
            {
                // Handle the error gracefully.
                Console.WriteLine($"Failed to insert OLE object with ProgId '{invalidProgId}'.");
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Save the document to the output folder.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OleObjectExample.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
