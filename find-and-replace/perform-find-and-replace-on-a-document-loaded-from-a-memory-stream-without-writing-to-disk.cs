using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a simple document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is the old value in the document.");

        // Save the document to a memory stream (no disk I/O).
        using MemoryStream stream = new MemoryStream();
        doc.Save(stream, SaveFormat.Docx);
        stream.Position = 0; // Reset for reading.

        // Load the document from the memory stream.
        Document loaded = new Document(stream);

        // Perform a find-and-replace operation.
        int replacedCount = loaded.Range.Replace("old", "new", new FindReplaceOptions());

        // Validate that a replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Prepare a simple report.
        var report = new
        {
            ReplacementCount = replacedCount,
            ResultText = loaded.GetText().Trim()
        };

        // Output the report as JSON.
        string json = JsonConvert.SerializeObject(report, Formatting.Indented);
        Console.WriteLine(json);
    }
}
