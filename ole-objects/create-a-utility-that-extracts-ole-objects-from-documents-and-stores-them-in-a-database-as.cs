using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class OleExtractor
{
    // Path to the source Word document. Adjust as needed.
    private const string InputDocumentPath = "input.docx";

    // Folder where extracted OLE objects will be saved (simulating storage).
    private const string OutputFolder = "ExtractedOleObjects";

    public static void Main()
    {
        // Ensure the output directory exists.
        Directory.CreateDirectory(OutputFolder);

        // Load the Word document. If the file does not exist, create an empty document instead.
        Document doc;
        if (File.Exists(InputDocumentPath))
        {
            doc = new Document(InputDocumentPath); // Load existing document
        }
        else
        {
            Console.WriteLine($"Input file \"{InputDocumentPath}\" not found. Creating an empty document.");
            doc = new Document(); // Create a blank document
        }

        // List to hold extracted OLE objects (file name + raw data).
        var extractedOleObjects = new List<(string FileName, byte[] Data)>();

        // Iterate over all shapes in the document.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            OleFormat ole = shape.OleFormat;
            if (ole == null)
                continue; // Not an OLE object.

            // Skip linked objects – they have no embedded data.
            if (ole.IsLink)
                continue;

            // Retrieve raw OLE data.
            byte[] rawData = ole.GetRawData();

            // Determine a file name for storage (use suggested name if available).
            string fileName = ole.SuggestedFileName;
            if (string.IsNullOrEmpty(fileName))
                fileName = $"OleObject_{Guid.NewGuid():N}.bin";

            // Save the raw data to a file (simulating a BLOB storage).
            string outputPath = Path.Combine(OutputFolder, fileName);
            File.WriteAllBytes(outputPath, rawData);

            // Keep the record in the in‑memory list.
            extractedOleObjects.Add((fileName, rawData));
        }

        // Simple console report.
        Console.WriteLine($"Extracted {extractedOleObjects.Count} OLE object(s) to folder \"{OutputFolder}\".");
    }
}
