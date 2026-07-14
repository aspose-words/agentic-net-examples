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

        // Dummy data to embed as an OLE object.
        byte[] dummyData = System.Text.Encoding.UTF8.GetBytes("Dummy OLE content");
        using (MemoryStream oleStream = new MemoryStream(dummyData))
        {
            // ProgId we intend to use for the OLE object.
            string progId = "Package";

            // Validate the ProgId before insertion.
            if (IsValidProgId(progId))
            {
                // Insert the OLE object into the document.
                // Parameters: stream, progId, asIcon (false), presentation (null).
                Shape oleShape = builder.InsertOleObject(oleStream, progId, false, null);

                // Optionally, set additional properties on the inserted OLE object.
                oleShape.OleFormat.IsLocked = false;
            }
            else
            {
                // If the ProgId is invalid, skip insertion or handle accordingly.
                Console.WriteLine($"Invalid ProgId: '{progId}'. OLE object not inserted.");
            }
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ValidatedOleObject.docx");
        doc.Save(outputPath);
    }

    // Simple validation method for ProgId strings.
    private static bool IsValidProgId(string progId)
    {
        // ProgId cannot be null, empty, or whitespace.
        if (string.IsNullOrWhiteSpace(progId))
            return false;

        // Example whitelist of known safe ProgIds.
        string[] allowedProgIds = { "Package", "Excel.Sheet", "Word.Document", "PowerPoint.Show" };

        // Return true if the ProgId is in the whitelist.
        foreach (string allowed in allowedProgIds)
        {
            if (string.Equals(progId, allowed, StringComparison.OrdinalIgnoreCase))
                return true;
        }

        // If not in whitelist, consider it invalid.
        return false;
    }
}
