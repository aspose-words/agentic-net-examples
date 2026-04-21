using System;
using System.IO;
using Aspose.Words;

public class MergeSplitDocuments
{
    public static void Main()
    {
        // Define a folder to store the sample and merged documents.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the split (sample) documents.
        string part1Path = Path.Combine(outputDir, "Part1.docx");
        string part2Path = Path.Combine(outputDir, "Part2.docx");

        // Create first sample document.
        Document part1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(part1);
        builder1.Writeln("This is the content of the first split document.");
        part1.Save(part1Path);

        // Create second sample document.
        Document part2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(part2);
        builder2.Writeln("This is the content of the second split document.");
        part2.Save(part2Path);

        // Load the split documents.
        Document srcDoc1 = new Document(part1Path);
        Document srcDoc2 = new Document(part2Path);

        // Create a destination document that will hold the merged result.
        Document mergedDoc = new Document();

        // Append the first document.
        mergedDoc.AppendDocument(srcDoc1, ImportFormatMode.KeepSourceFormatting);
        // Append the second document.
        mergedDoc.AppendDocument(srcDoc2, ImportFormatMode.KeepSourceFormatting);

        // Save the merged document.
        string mergedPath = Path.Combine(outputDir, "Merged.docx");
        mergedDoc.Save(mergedPath);

        // Validate that the merged file was created.
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException("Merged document was not created.");

        // Optional: indicate success.
        Console.WriteLine("Documents merged successfully. Output file: " + mergedPath);
    }
}
