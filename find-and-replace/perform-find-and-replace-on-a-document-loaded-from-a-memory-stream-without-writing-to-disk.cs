using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document in memory.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is the old value.");
        builder.Writeln("Another line with old value inside.");

        // Save the document to a memory stream.
        using MemoryStream sourceStream = new MemoryStream();
        sourceDoc.Save(sourceStream, SaveFormat.Docx);
        sourceStream.Position = 0; // Reset for reading.

        // Load the document from the memory stream.
        Document loadedDoc = new Document(sourceStream);

        // Perform find-and-replace.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = loadedDoc.Range.Replace("old", "new", options);

        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document to another memory stream (no disk I/O).
        using MemoryStream resultStream = new MemoryStream();
        loadedDoc.Save(resultStream, SaveFormat.Docx);
        resultStream.Position = 0; // Reset if further processing is needed.

        // Output a simple verification to the console.
        Console.WriteLine($"Replacements performed: {replacedCount}");
        Console.WriteLine("Modified document text:");
        Console.WriteLine(loadedDoc.GetText().Trim());
    }
}
