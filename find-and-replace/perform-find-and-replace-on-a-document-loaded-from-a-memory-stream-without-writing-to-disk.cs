using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required by the task specification
using Newtonsoft.Json; // Required by the task specification

public class Program
{
    public static void Main()
    {
        // Create a sample document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello old value. This is a test of old value replacement.");

        // Save the document to a memory stream.
        using MemoryStream inputStream = new MemoryStream();
        doc.Save(inputStream, SaveFormat.Docx);
        inputStream.Position = 0; // Reset stream position for reading.

        // Load the document from the memory stream.
        Document loadedDoc = new Document(inputStream);

        // Perform a find-and-replace operation.
        FindReplaceOptions options = new FindReplaceOptions();
        int replaceCount = loadedDoc.Range.Replace("old value", "new value", options);
        if (replaceCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Output the resulting text to the console.
        Console.WriteLine(loadedDoc.GetText());

        // Save the modified document to another memory stream (no disk I/O).
        using MemoryStream outputStream = new MemoryStream();
        loadedDoc.Save(outputStream, SaveFormat.Docx);
        // The outputStream now contains the updated document.
    }
}
