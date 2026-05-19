using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    // Validates that the provided ProgID is not null, empty, or whitespace.
    private static bool IsValidProgId(string progId)
    {
        return !string.IsNullOrWhiteSpace(progId);
    }

    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample data to embed as an OLE package.
        byte[] sampleData = System.Text.Encoding.UTF8.GetBytes("This is sample OLE package content.");
        using (MemoryStream dataStream = new MemoryStream(sampleData))
        {
            string progId = "Package"; // ProgID for a generic OLE package.

            // Validate the ProgID before attempting insertion.
            if (IsValidProgId(progId))
            {
                // Insert the OLE object using the validated ProgID.
                // Parameters: stream, progId, asIcon (false), presentation (null).
                Shape oleShape = builder.InsertOleObject(dataStream, progId, false, null);

                // Optionally, set a display name for the package.
                if (oleShape?.OleFormat?.OlePackage != null)
                {
                    oleShape.OleFormat.OlePackage.FileName = "Sample.txt";
                    oleShape.OleFormat.OlePackage.DisplayName = "Sample.txt";
                }
            }
            else
            {
                // If the ProgID is invalid, skip insertion (could log or handle as needed).
                Console.WriteLine("Invalid ProgID provided. OLE object insertion skipped.");
            }
        }

        // Save the resulting document.
        doc.Save("ValidatedOle.docx");
    }
}
