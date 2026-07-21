using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create the first document with some content.
        Document doc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(doc1);
        builder1.Writeln("Alpha");

        // Save the first document to a memory stream.
        using MemoryStream ms1 = new MemoryStream();
        doc1.Save(ms1, SaveFormat.Docx);
        ms1.Position = 0; // Reset the stream position for reading.

        // Create the second document with different content.
        Document doc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(doc2);
        builder2.Writeln("Beta");

        // Save the second document to a memory stream.
        using MemoryStream ms2 = new MemoryStream();
        doc2.Save(ms2, SaveFormat.Docx);
        ms2.Position = 0; // Reset the stream position for reading.

        // Load the documents back from the memory streams.
        Document loaded1 = new Document(ms1);
        Document loaded2 = new Document(ms2);

        // Perform the comparison. The revisions will be added to loaded1.
        loaded1.Compare(loaded2, "Author", DateTime.Now);

        // Verify that at least one revision was created.
        if (loaded1.Revisions.Count == 0)
            throw new InvalidOperationException("Expected at least one revision after comparison.");

        // Save the comparison result to a memory stream.
        using MemoryStream resultStream = new MemoryStream();
        loaded1.Save(resultStream, SaveFormat.Docx);
        byte[] resultBytes = resultStream.ToArray(); // The resulting document as a byte array.

        // (Optional) Write the byte array to a file for manual inspection.
        File.WriteAllBytes("ComparisonResult.docx", resultBytes);
    }
}
