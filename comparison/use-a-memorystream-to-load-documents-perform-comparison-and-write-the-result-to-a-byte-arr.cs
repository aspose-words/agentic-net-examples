using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the first document with distinct content.
        Document doc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(doc1);
        builder1.Writeln("Alpha");

        // Save the first document into a memory stream.
        using MemoryStream ms1 = new MemoryStream();
        doc1.Save(ms1, SaveFormat.Docx);
        ms1.Position = 0;

        // Create the second document with different content.
        Document doc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(doc2);
        builder2.Writeln("Beta");

        // Save the second document into a memory stream.
        using MemoryStream ms2 = new MemoryStream();
        doc2.Save(ms2, SaveFormat.Docx);
        ms2.Position = 0;

        // Load documents from the memory streams.
        Document loaded1 = new Document(ms1);
        Document loaded2 = new Document(ms2);

        // Perform comparison between the two documents.
        loaded1.Compare(loaded2, "Comparer", DateTime.Now);

        // Ensure that revisions were generated.
        if (loaded1.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Save the comparison result to a memory stream and obtain a byte array.
        using MemoryStream resultStream = new MemoryStream();
        loaded1.Save(resultStream, SaveFormat.Docx);
        byte[] resultBytes = resultStream.ToArray();

        // Output the size of the resulting byte array.
        Console.WriteLine($"Comparison result byte array length: {resultBytes.Length}");
    }
}
