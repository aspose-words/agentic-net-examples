using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class StreamComparisonExample
{
    public static void Main()
    {
        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This is the original document.");

        // Save the original document to a memory stream.
        using MemoryStream originalStream = new MemoryStream();
        original.Save(originalStream, SaveFormat.Docx);
        originalStream.Position = 0; // Reset stream position for reading.

        // Create the revised document with different content.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This is the revised document with changes.");

        // Save the revised document to a memory stream.
        using MemoryStream revisedStream = new MemoryStream();
        revised.Save(revisedStream, SaveFormat.Docx);
        revisedStream.Position = 0; // Reset stream position for reading.

        // Load the documents back from the memory streams.
        Document loadedOriginal = new Document(originalStream);
        Document loadedRevised = new Document(revisedStream);

        // Perform comparison between the two documents.
        loadedOriginal.Compare(loadedRevised, "Comparer", DateTime.Now);

        // Verify that at least one revision was created.
        if (loadedOriginal.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Save the comparison result to a new memory stream.
        using MemoryStream resultStream = new MemoryStream();
        loadedOriginal.Save(resultStream, SaveFormat.Docx);

        // Convert the result stream to a byte array.
        byte[] resultBytes = resultStream.ToArray();

        // Output the size of the resulting byte array to the console.
        Console.WriteLine($"Comparison result byte array length: {resultBytes.Length}");
    }
}
