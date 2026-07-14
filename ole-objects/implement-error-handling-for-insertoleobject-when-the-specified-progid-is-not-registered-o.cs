using System;
using System.IO;
using System.Runtime.InteropServices;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare dummy data to embed as an OLE object.
        byte[] dummyData = System.Text.Encoding.UTF8.GetBytes("Dummy content");
        using (MemoryStream dataStream = new MemoryStream(dummyData))
        {
            // Attempt to insert an OLE object with a ProgId that is unlikely to be registered.
            // Wrap the call in a try-catch block to handle the error gracefully.
            try
            {
                // The ProgId "NonExistent.ProgId" does not correspond to any installed application.
                builder.InsertOleObject(dataStream, "NonExistent.ProgId", false, null);
                Console.WriteLine("OLE object inserted successfully.");
            }
            catch (COMException comEx)
            {
                // COMException is thrown when the ProgId cannot be resolved.
                Console.WriteLine($"COMException caught: {comEx.Message}");
            }
            catch (Exception ex)
            {
                // Catch any other unexpected exceptions.
                Console.WriteLine($"Exception caught: {ex.Message}");
            }
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
