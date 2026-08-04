using System;
using System.IO;
using Aspose.Words;

public class MergeSplitDocuments
{
    public static void Main()
    {
        // Define a folder for the sample and output files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "MergedDocs");
        Directory.CreateDirectory(outputDir);

        // Paths for the split documents.
        string part1Path = Path.Combine(outputDir, "Part1.docx");
        string part2Path = Path.Combine(outputDir, "Part2.docx");
        string mergedPath = Path.Combine(outputDir, "Merged.docx");

        // -----------------------------------------------------------------
        // Create sample split documents.
        // -----------------------------------------------------------------
        CreateSampleDocument(part1Path, "This is the content of the first split document.");
        CreateSampleDocument(part2Path, "This is the content of the second split document.");

        // -----------------------------------------------------------------
        // Load the split documents.
        // -----------------------------------------------------------------
        Document part1 = new Document(part1Path);
        Document part2 = new Document(part2Path);

        // -----------------------------------------------------------------
        // Merge the loaded documents using AppendDocument.
        // -----------------------------------------------------------------
        Document merged = new Document(); // Starts with a single empty section.
        merged.AppendDocument(part1, ImportFormatMode.KeepSourceFormatting);
        merged.AppendDocument(part2, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document.
        merged.Save(mergedPath);

        // Simple validation to ensure the merged file was created.
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException("Merged document was not created.");

        // Optional: verify that the merged document contains text from both parts.
        string mergedText = merged.GetText();
        if (!mergedText.Contains("first split document") || !mergedText.Contains("second split document"))
            throw new InvalidOperationException("Merged document does not contain expected content.");
    }

    // Helper method to create a simple document with a single paragraph of text.
    private static void CreateSampleDocument(string filePath, string paragraphText)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(paragraphText);
        doc.Save(filePath);
    }
}
