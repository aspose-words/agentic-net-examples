using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document in memory.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is the old value that will be replaced.");
        builder.Writeln("Another line with old text.");

        // Save the document to a MemoryStream (no disk I/O).
        using MemoryStream sourceStream = new MemoryStream();
        sourceDoc.Save(sourceStream, SaveFormat.Docx);
        sourceStream.Position = 0; // Reset for reading.

        // Load the document from the MemoryStream.
        Document loadedDoc = new Document(sourceStream);

        // Perform a find-and-replace operation.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = loadedDoc.Range.Replace("old", "new", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document to another MemoryStream (still no disk I/O).
        using MemoryStream resultStream = new MemoryStream();
        loadedDoc.Save(resultStream, SaveFormat.Docx);

        // Output simple verification information.
        Console.WriteLine($"Replacements performed: {replacedCount}");
        Console.WriteLine($"Modified document size (bytes): {resultStream.Length}");
    }
}
