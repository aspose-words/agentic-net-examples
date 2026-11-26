using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Saving;
using Aspose.Drawing; // Required by the task specification
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello old value. This is another old value.");

        // Save the document to a MemoryStream (no disk I/O).
        using MemoryStream sourceStream = new MemoryStream();
        doc.Save(sourceStream, SaveFormat.Docx);
        sourceStream.Position = 0; // Reset for reading.

        // Load the document from the MemoryStream.
        Document loadedDoc = new Document(sourceStream);

        // Perform a plain‑text find‑and‑replace.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = loadedDoc.Range.Replace("old", "new", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Optionally, save the modified document to another MemoryStream to prove it was changed.
        using MemoryStream outputStream = new MemoryStream();
        loadedDoc.Save(outputStream, SaveFormat.Docx);
        if (outputStream.Length == 0)
            throw new InvalidOperationException("The output stream is empty after saving the modified document.");

        // Create a simple report of the operation.
        var report = new
        {
            OriginalText = "Hello old value. This is another old value.",
            SearchTerm = "old",
            Replacement = "new",
            ReplacementsMade = replacedCount,
            // Demonstrate usage of Aspose.Drawing (e.g., a color constant) as required by the task.
            HighlightColor = Color.Red.ToArgb()
        };

        // Serialize the report to JSON and write it to the console.
        string jsonReport = JsonConvert.SerializeObject(report, Formatting.Indented);
        Console.WriteLine(jsonReport);
    }
}
