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

        // Define the ProgID we intend to use for the OLE object.
        string progId = "Package"; // Example of a known ProgID.

        // Validate the ProgID before insertion.
        if (!IsProgIdValid(progId))
        {
            // If the ProgID is not valid, skip insertion and finish.
            Console.WriteLine($"ProgID \"{progId}\" is not valid. Skipping OLE insertion.");
            doc.Save("ValidatedOleObject.docx");
            return;
        }

        // Prepare dummy data for the OLE object (e.g., a simple byte array).
        byte[] dummyData = new byte[] { 0x50, 0x4B, 0x03, 0x04 }; // Beginning of a ZIP file header.
        using (MemoryStream stream = new MemoryStream(dummyData))
        {
            // Insert the OLE object using the validated ProgID.
            // Parameters: stream, progId, asIcon (false), presentation (null).
            Shape oleShape = builder.InsertOleObject(stream, progId, false, null);

            // Optional: verify that the inserted object's ProgID matches the expected value.
            string insertedProgId = oleShape.OleFormat.ProgId;
            Console.WriteLine($"Inserted OLE object ProgID: {insertedProgId}");
        }

        // Save the document to the file system.
        doc.Save("ValidatedOleObject.docx");
    }

    // Simple validation method for ProgID strings.
    private static bool IsProgIdValid(string progId)
    {
        // ProgID must not be null or empty.
        if (string.IsNullOrEmpty(progId))
            return false;

        // Example whitelist of known safe ProgIDs.
        string[] allowedProgIds = new string[]
        {
            "Package",          // Generic OLE package.
            "Excel.Sheet",      // Microsoft Excel.
            "Word.Document",    // Microsoft Word.
            "PowerPoint.Show",  // Microsoft PowerPoint.
            "Visio.Drawing"     // Microsoft Visio.
        };

        // Check if the provided ProgID is in the whitelist.
        foreach (string allowed in allowedProgIds)
        {
            if (string.Equals(progId, allowed, StringComparison.OrdinalIgnoreCase))
                return true;
        }

        // If not found in the whitelist, consider it invalid.
        return false;
    }
}
