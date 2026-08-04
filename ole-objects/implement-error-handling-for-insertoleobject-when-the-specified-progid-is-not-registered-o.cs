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

        // Prepare a simple memory stream as the OLE data source.
        byte[] dummyData = System.Text.Encoding.UTF8.GetBytes("Dummy OLE data");
        using (MemoryStream stream = new MemoryStream(dummyData))
        {
            builder.Writeln("Attempting to insert OLE object with an invalid ProgId:");

            try
            {
                // Use a ProgId that is not registered on the system to trigger an exception.
                builder.InsertOleObject(stream, "NonExistent.ProgId", false, null);
                builder.Writeln("OLE object inserted successfully.");
            }
            catch (Exception ex)
            {
                // Gracefully handle the error and inform the user.
                builder.Writeln($"Failed to insert OLE object: {ex.Message}");
            }
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OleObjectErrorHandling.docx");
        doc.Save(outputPath);
    }
}
