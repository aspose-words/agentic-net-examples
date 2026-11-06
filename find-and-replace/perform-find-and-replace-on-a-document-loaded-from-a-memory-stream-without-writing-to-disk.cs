using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a simple document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample text with placeholder that will be replaced.");

        // Save the document to a memory stream (no disk I/O).
        using MemoryStream inputStream = new MemoryStream();
        doc.Save(inputStream, SaveFormat.Docx);
        inputStream.Position = 0; // Reset stream for reading.

        // Load the document from the memory stream.
        Document loadedDoc = new Document(inputStream);

        // Perform a find-and-replace operation.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = loadedDoc.Range.Replace("placeholder", "value", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document to another memory stream.
        using MemoryStream outputStream = new MemoryStream();
        loadedDoc.Save(outputStream, SaveFormat.Docx);

        // Output simple verification information.
        Console.WriteLine($"Replacements made: {replacedCount}");
        Console.WriteLine($"Modified document size: {outputStream.Length} bytes");
    }
}
