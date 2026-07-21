using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class InsertOleObjectExample
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare some dummy data to embed as an OLE object.
        byte[] dummyData = new byte[] { 0x00, 0x01, 0x02, 0x03 };
        using (MemoryStream oleStream = new MemoryStream(dummyData))
        {
            // Use a ProgId that is unlikely to be registered on the system.
            string invalidProgId = "NonExistent.ProgId";

            try
            {
                // Attempt to insert the OLE object. This may throw if the ProgId is not registered.
                builder.InsertOleObject(oleStream, invalidProgId, asIcon: false, presentation: null);
                Console.WriteLine("OLE object inserted successfully.");
            }
            catch (Exception ex)
            {
                // Handle the error gracefully and inform the user.
                Console.WriteLine($"Failed to insert OLE object with ProgId '{invalidProgId}'.");
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "InsertOleObjectResult.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
