using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define a folder for all generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create sample split documents that we will later merge.
        // -----------------------------------------------------------------
        string part1Path = Path.Combine(outputDir, "Part1.docx");
        string part2Path = Path.Combine(outputDir, "Part2.docx");

        // Document 1
        Document part1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(part1);
        builder1.Writeln("This is the first split document.");
        part1.Save(part1Path);

        // Document 2
        Document part2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(part2);
        builder2.Writeln("This is the second split document.");
        part2.Save(part2Path);

        // -----------------------------------------------------------------
        // 2. Load the split documents.
        // -----------------------------------------------------------------
        Document srcDoc1 = new Document(part1Path);
        Document srcDoc2 = new Document(part2Path);

        // -----------------------------------------------------------------
        // 3. Merge them using AppendDocument.
        // -----------------------------------------------------------------
        Document mergedDoc = new Document(); // starts with a single empty section
        mergedDoc.AppendDocument(srcDoc1, ImportFormatMode.KeepSourceFormatting);
        mergedDoc.AppendDocument(srcDoc2, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document.
        string mergedPath = Path.Combine(outputDir, "Combined.docx");
        mergedDoc.Save(mergedPath);

        // -----------------------------------------------------------------
        // 4. Simple validation – ensure the combined file exists and contains both texts.
        // -----------------------------------------------------------------
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException("Merged document was not created.");

        Document verifyDoc = new Document(mergedPath);
        string mergedText = verifyDoc.GetText();

        if (!mergedText.Contains("first split document") || !mergedText.Contains("second split document"))
            throw new InvalidOperationException("Merged document does not contain expected content.");

        // All done – the program will exit automatically.
    }
}
