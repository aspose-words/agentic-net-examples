using System;
using System.IO;
using Aspose.Words;

public class MergeSplitDocuments
{
    public static void Main()
    {
        // Define folder for temporary and output files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the sample split documents.
        string part1Path = Path.Combine(outputDir, "Part1.docx");
        string part2Path = Path.Combine(outputDir, "Part2.docx");
        string combinedPath = Path.Combine(outputDir, "Combined.docx");

        // -----------------------------------------------------------------
        // Create sample split documents (normally these would be the result
        // of a previous split operation). Each document contains simple text.
        // -----------------------------------------------------------------
        Document part1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(part1);
        builder1.Writeln("This is the content of the first split document.");
        part1.Save(part1Path);

        Document part2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(part2);
        builder2.Writeln("This is the content of the second split document.");
        part2.Save(part2Path);

        // -----------------------------------------------------------------
        // Load the split documents that need to be merged.
        // -----------------------------------------------------------------
        Document srcDoc1 = new Document(part1Path);
        Document srcDoc2 = new Document(part2Path);

        // -----------------------------------------------------------------
        // Create the destination document and append the loaded documents.
        // KeepSourceFormatting preserves the original formatting of each part.
        // -----------------------------------------------------------------
        Document dstDoc = new Document();
        dstDoc.AppendDocument(srcDoc1, ImportFormatMode.KeepSourceFormatting);
        dstDoc.AppendDocument(srcDoc2, ImportFormatMode.KeepSourceFormatting);

        // Save the merged document.
        dstDoc.Save(combinedPath);

        // -----------------------------------------------------------------
        // Validation: ensure the combined file exists and contains both parts.
        // -----------------------------------------------------------------
        if (!File.Exists(combinedPath))
            throw new Exception("Merged document was not created.");

        Document verificationDoc = new Document(combinedPath);
        string combinedText = verificationDoc.GetText();

        if (!combinedText.Contains("first split document") || !combinedText.Contains("second split document"))
            throw new Exception("Merged document does not contain expected content.");

        // The program finishes without requiring any user interaction.
    }
}
