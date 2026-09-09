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

        // Dummy OLE data – in a real scenario this would be the actual file bytes.
        byte[] dummyData = new byte[] { 0x00 };
        using (MemoryStream oleStream = new MemoryStream(dummyData))
        {
            // The ProgID we intend to use for the OLE object.
            string progId = "Package";

            // Validate the ProgID before attempting insertion.
            if (IsProgIdValid(progId))
            {
                // Insert the OLE object using the validated ProgID.
                Shape oleShape = builder.InsertOleObject(oleStream, progId, false, null);

                // Retrieve and display the ProgID of the inserted object.
                string insertedProgId = oleShape.OleFormat.ProgId;
                Console.WriteLine($"Inserted OLE object with ProgId: {insertedProgId}");
            }
            else
            {
                Console.WriteLine($"ProgId '{progId}' is not valid. Insertion skipped.");
            }
        }

        // Save the document to the file system.
        doc.Save("ValidatedOleObject.docx");
    }

    // Simple validation logic for a ProgID.
    private static bool IsProgIdValid(string progId)
    {
        // ProgId must be non‑null, non‑empty and must not contain whitespace.
        return !string.IsNullOrEmpty(progId) && !progId.Contains(" ");
    }
}
