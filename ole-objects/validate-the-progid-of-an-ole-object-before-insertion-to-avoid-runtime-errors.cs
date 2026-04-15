using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class ValidateOleProgIdExample
{
    public static void Main()
    {
        // Path where the resulting document will be saved.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ValidatedOleObject.docx");

        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Example data to embed as an OLE object (a simple text file content).
        byte[] dummyData = System.Text.Encoding.UTF8.GetBytes("Sample OLE package content");
        using (MemoryStream oleStream = new MemoryStream(dummyData))
        {
            // ProgId we intend to use for the OLE object.
            string progId = "Package";

            // Validate the ProgId before insertion.
            if (IsValidProgId(progId))
            {
                // Insert the OLE object using the validated ProgId.
                // Parameters: stream, progId, asIcon (false = display content), presentation (null = default icon).
                Shape oleShape = builder.InsertOleObject(oleStream, progId, false, null);

                // Optionally, set a display name for the OLE package.
                if (oleShape?.OleFormat?.OlePackage != null)
                {
                    oleShape.OleFormat.OlePackage.FileName = "SamplePackage.bin";
                    oleShape.OleFormat.OlePackage.DisplayName = "Sample Package";
                }
            }
            else
            {
                // If the ProgId is invalid, write a message to the console (no user interaction required).
                Console.WriteLine($"ProgId \"{progId}\" is invalid. OLE object was not inserted.");
            }
        }

        // Save the document to disk.
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }

    // Simple validation: ProgId must be non‑empty and contain at least one dot (e.g., "Excel.Sheet").
    private static bool IsValidProgId(string progId)
    {
        return !string.IsNullOrEmpty(progId) && progId.Contains(".");
    }
}
