using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Saving;
using Aspose.Drawing; // Required package, not used directly
using Newtonsoft.Json; // Required package, not used directly

public class Program
{
    public static void Main()
    {
        // Create a sample document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello ReplaceMe world.");
        builder.Writeln("ReplaceMe appears twice: ReplaceMe.");

        // Save the document to a memory stream (no disk I/O).
        using MemoryStream inputStream = new MemoryStream();
        doc.Save(inputStream, SaveFormat.Docx);
        inputStream.Position = 0; // Reset for reading.

        // Load the document from the memory stream.
        Document loadedDoc = new Document(inputStream);

        // Perform a find-and-replace operation.
        FindReplaceOptions options = new FindReplaceOptions();
        int replaceCount = loadedDoc.Range.Replace("ReplaceMe", "Updated", options);

        // Validate that at least one replacement occurred.
        if (replaceCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document to another memory stream.
        using MemoryStream outputStream = new MemoryStream();
        loadedDoc.Save(outputStream, SaveFormat.Docx);

        // Output simple verification information.
        Console.WriteLine($"Replacements made: {replaceCount}");
        Console.WriteLine($"Resulting document size: {outputStream.Length} bytes");
    }
}
