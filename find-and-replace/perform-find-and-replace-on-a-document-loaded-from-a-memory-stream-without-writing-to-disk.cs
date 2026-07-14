using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document containing the word old.");

        // Save the document to a memory stream (no disk I/O).
        using MemoryStream inputStream = new MemoryStream();
        doc.Save(inputStream, SaveFormat.Docx);
        inputStream.Position = 0; // Reset for reading.

        // Load the document from the memory stream.
        Document loadedDoc = new Document(inputStream);

        // Perform a find-and-replace operation.
        FindReplaceOptions options = new FindReplaceOptions();
        int replaceCount = loadedDoc.Range.Replace("old", "new", options);

        // Validate that a replacement occurred.
        if (replaceCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Optionally, output the resulting text to the console.
        Console.WriteLine("Modified document text:");
        Console.WriteLine(loadedDoc.GetText());

        // Save the modified document to another memory stream (still no disk I/O).
        using MemoryStream outputStream = new MemoryStream();
        loadedDoc.Save(outputStream, SaveFormat.Docx);
        // The outputStream now contains the updated DOCX file.
    }
}
