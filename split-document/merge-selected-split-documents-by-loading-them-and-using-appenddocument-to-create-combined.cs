using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define output folder and ensure it exists
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputFolder);

        // Paths for the split documents and the merged result
        string doc1Path = Path.Combine(outputFolder, "Part1.docx");
        string doc2Path = Path.Combine(outputFolder, "Part2.docx");
        string mergedPath = Path.Combine(outputFolder, "Merged.docx");

        // Create first sample document
        Document doc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(doc1);
        builder1.Writeln("This is the content of the first split document.");
        doc1.Save(doc1Path);

        // Create second sample document
        Document doc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(doc2);
        builder2.Writeln("This is the content of the second split document.");
        doc2.Save(doc2Path);

        // Load the split documents
        Document part1 = new Document(doc1Path);
        Document part2 = new Document(doc2Path);

        // Append the second document to the first one, preserving source formatting
        part1.AppendDocument(part2, ImportFormatMode.KeepSourceFormatting);

        // Save the merged document
        part1.Save(mergedPath);

        // Validation: ensure merged file exists and contains both pieces of text
        if (!File.Exists(mergedPath))
            throw new FileNotFoundException("Merged document was not created.", mergedPath);

        Document mergedDoc = new Document(mergedPath);
        string mergedText = mergedDoc.GetText();

        if (!mergedText.Contains("first split document") || !mergedText.Contains("second split document"))
            throw new InvalidOperationException("Merged document does not contain expected content.");

        // Program completed successfully
    }
}
