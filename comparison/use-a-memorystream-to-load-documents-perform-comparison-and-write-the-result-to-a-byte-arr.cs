using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create the original document in memory.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Alpha");

        // Save the original document to a MemoryStream.
        using MemoryStream msOriginal = new MemoryStream();
        original.Save(msOriginal, SaveFormat.Docx);
        msOriginal.Position = 0; // Reset for reading.

        // Create the revised document in memory.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Beta");

        // Save the revised document to a MemoryStream.
        using MemoryStream msRevised = new MemoryStream();
        revised.Save(msRevised, SaveFormat.Docx);
        msRevised.Position = 0; // Reset for reading.

        // Load the documents back from the streams.
        Document loadedOriginal = new Document(msOriginal);
        Document loadedRevised = new Document(msRevised);

        // Perform the comparison.
        loadedOriginal.Compare(loadedRevised, "Author", DateTime.Now);

        // Verify that at least one revision was created.
        if (loadedOriginal.Revisions.Count == 0)
            throw new InvalidOperationException("Expected at least one revision after comparison.");

        // Save the comparison result to a MemoryStream and obtain the byte array.
        using MemoryStream resultStream = new MemoryStream();
        loadedOriginal.Save(resultStream, SaveFormat.Docx);
        byte[] resultBytes = resultStream.ToArray();

        // Output the size of the resulting byte array to confirm execution.
        Console.WriteLine($"Comparison result byte array length: {resultBytes.Length}");
    }
}
