using System;
using System.IO;
using Aspose.Words;

public class MergeSplitDocuments
{
    public static void Main()
    {
        // Define a folder to store the sample split documents and the merged result.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);

        // Create three sample documents that will act as previously split parts.
        CreateSampleDocument(Path.Combine(dataDir, "Part1.docx"), "This is the content of part 1.");
        CreateSampleDocument(Path.Combine(dataDir, "Part2.docx"), "This is the content of part 2.");
        CreateSampleDocument(Path.Combine(dataDir, "Part3.docx"), "This is the content of part 3.");

        // Prepare the destination document that will receive the appended parts.
        Document mergedDoc = new Document();

        // Load each split document and append it to the destination.
        foreach (string partPath in new[] {
                     Path.Combine(dataDir, "Part1.docx"),
                     Path.Combine(dataDir, "Part2.docx"),
                     Path.Combine(dataDir, "Part3.docx") })
        {
            Document partDoc = new Document(partPath);
            mergedDoc.AppendDocument(partDoc, ImportFormatMode.KeepSourceFormatting);
        }

        // Save the combined document.
        string mergedPath = Path.Combine(dataDir, "Merged.docx");
        mergedDoc.Save(mergedPath);

        // Simple validation – ensure the merged file was created.
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException("Merged document was not created.");

        // Optionally, output the location of the merged file.
        Console.WriteLine($"Merged document saved to: {mergedPath}");
    }

    // Helper method to create a simple document with a single paragraph of text.
    private static void CreateSampleDocument(string filePath, string text)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(text);
        doc.Save(filePath);
    }
}
