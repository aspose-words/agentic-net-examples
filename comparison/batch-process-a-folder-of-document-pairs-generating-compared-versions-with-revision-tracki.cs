using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Set up folders for input documents and output comparison results.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "ComparisonInput");
        string outputDir = Path.Combine(baseDir, "ComparisonOutput");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Number of document pairs to process.
        const int pairCount = 3;

        // Create sample document pairs with deterministic differences.
        for (int i = 1; i <= pairCount; i++)
        {
            // Original document.
            var original = new Document();
            var builderOrig = new DocumentBuilder(original);
            builderOrig.Writeln($"Document {i} original content.");
            builderOrig.Writeln("Common line.");
            string originalPath = Path.Combine(inputDir, $"doc{i}_original.docx");
            original.Save(originalPath);

            // Revised document with changes.
            var revised = new Document();
            var builderRev = new DocumentBuilder(revised);
            builderRev.Writeln($"Document {i} revised content."); // changed line
            builderRev.Writeln("Common line.");                    // unchanged line
            builderRev.Writeln("Additional line.");               // extra line
            string revisedPath = Path.Combine(inputDir, $"doc{i}_revised.docx");
            revised.Save(revisedPath);
        }

        // Batch compare each pair and save the result with revision tracking.
        for (int i = 1; i <= pairCount; i++)
        {
            string originalPath = Path.Combine(inputDir, $"doc{i}_original.docx");
            string revisedPath = Path.Combine(inputDir, $"doc{i}_revised.docx");

            // Load the documents.
            var originalDoc = new Document(originalPath);
            var revisedDoc = new Document(revisedPath);

            // Perform comparison; revisions are added to the original document.
            originalDoc.Compare(revisedDoc, "BatchUser", DateTime.Now);

            // Report the number of revisions detected for this pair.
            int revisionCount = originalDoc.Revisions.Count;
            Console.WriteLine($"Pair {i}: {revisionCount} revision(s) detected.");

            // Save the compared document preserving revisions.
            string resultPath = Path.Combine(outputDir, $"doc{i}_compared.docx");
            originalDoc.Save(resultPath);
        }
    }
}
