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

        // Prepare a simple byte array to act as the OLE object's data stream.
        byte[] dummyData = new byte[] { 0x00, 0x01, 0x02, 0x03 };
        using (MemoryStream dataStream = new MemoryStream(dummyData))
        {
            try
            {
                // Attempt to insert an OLE object with a ProgId that is unlikely to be registered.
                // This should raise an exception on systems where the ProgId does not exist.
                builder.InsertOleObject(dataStream, "Invalid.ProgId", false, null);
                Console.WriteLine("OLE object inserted successfully.");
            }
            catch (Exception ex)
            {
                // Handle the error gracefully and inform the user.
                Console.WriteLine($"Failed to insert OLE object with ProgId 'Invalid.ProgId': {ex.Message}");

                // Fallback: insert the same data as a generic package OLE object.
                // Reset the stream position before reusing it.
                dataStream.Position = 0;
                try
                {
                    Shape fallbackShape = builder.InsertOleObject(dataStream, "Package", true, null);
                    // Optionally set a display name for the package.
                    fallbackShape.OleFormat.OlePackage.FileName = "fallback.bin";
                    fallbackShape.OleFormat.OlePackage.DisplayName = "Fallback Package";
                    Console.WriteLine("Inserted fallback OLE package instead.");
                }
                catch (Exception fallbackEx)
                {
                    Console.WriteLine($"Fallback insertion also failed: {fallbackEx.Message}");
                }
            }
        }

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OleObjectExample.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
