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

        // Path to a sample file that could be embedded (not required for the error case).
        string sampleFilePath = "sample.txt";

        // Ensure the sample file exists to avoid file‑not‑found errors.
        File.WriteAllText(sampleFilePath, "This is a sample text file.");

        try
        {
            // Open the file as a stream because the InsertOleObject overload expects streams.
            using (FileStream oleStream = File.OpenRead(sampleFilePath))
            {
                // Attempt to insert an OLE object using a ProgId that is not registered.
                // This will throw an exception because the ProgId does not exist on the system.
                // The fourth parameter (icon stream) is set to null because we are not providing a custom icon.
                builder.InsertOleObject(oleStream, "NonExistent.ProgId", false, null);
            }
        }
        catch (Exception ex)
        {
            // Handle the error gracefully and inform the user.
            Console.WriteLine("Error inserting OLE object: " + ex.Message);
        }

        // Save the document (it will be empty or contain whatever succeeded before the error).
        doc.Save("Output.docx");
    }
}
