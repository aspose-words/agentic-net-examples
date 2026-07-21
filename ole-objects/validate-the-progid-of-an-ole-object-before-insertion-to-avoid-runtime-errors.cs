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

        // Dummy data to embed as an OLE object (e.g., a simple text file content).
        byte[] dummyData = System.Text.Encoding.UTF8.GetBytes("Sample OLE content");
        using (MemoryStream oleStream = new MemoryStream(dummyData))
        {
            // Define the ProgID we intend to use.
            string progId = "Package"; // Example ProgID for a generic OLE package.

            // Validate the ProgID before attempting insertion.
            if (IsValidProgId(progId))
            {
                // Insert the OLE object as an embedded object (not as an icon).
                // Presentation stream is null, so Aspose.Words will use a default icon if needed.
                Shape oleShape = builder.InsertOleObject(oleStream, progId, asIcon: false, presentation: null);

                // After insertion, we can read back the ProgID to confirm it was set correctly.
                string insertedProgId = oleShape.OleFormat.ProgId;
                Console.WriteLine($"OLE object inserted with ProgID: {insertedProgId}");
            }
            else
            {
                Console.WriteLine($"Invalid ProgID '{progId}'. OLE object was not inserted.");
            }
        }

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ValidatedOleObject.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }

    // Simple validation: ProgID must be non‑empty and contain at least one dot (e.g., "Excel.Sheet").
    private static bool IsValidProgId(string progId)
    {
        if (string.IsNullOrEmpty(progId))
            return false;

        // Basic pattern check – most ProgIDs contain a period separating the application name and type.
        return progId.Contains(".");
    }
}
