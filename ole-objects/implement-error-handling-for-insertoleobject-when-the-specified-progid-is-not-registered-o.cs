using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to work with it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare some dummy data to act as the OLE object's content.
        byte[] dummyData = System.Text.Encoding.UTF8.GetBytes("Dummy OLE content");
        using (MemoryStream stream = new MemoryStream(dummyData))
        {
            // Use an intentionally invalid ProgId to trigger an error.
            string progId = "NonExistent.ProgId";

            try
            {
                // Attempt to insert the OLE object.
                builder.InsertOleObject(stream, progId, false, null);
                Console.WriteLine("OLE object inserted successfully.");
            }
            catch (Exception ex)
            {
                // Handle the case where the ProgId is not registered.
                Console.WriteLine($"Failed to insert OLE object. ProgId '{progId}' may not be registered.");
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OleObjectDemo.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
