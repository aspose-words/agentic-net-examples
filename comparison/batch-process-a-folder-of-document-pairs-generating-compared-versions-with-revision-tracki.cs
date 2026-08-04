using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Comparing;

public class BatchDocumentComparison
{
    public static void Main()
    {
        // Define input and output folders relative to the current directory.
        string inputDir = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonInput");
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonOutput");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Seed the input folder with a few pairs of sample documents.
        // Each pair consists of a version A and a version B with intentional differences.
        for (int i = 1; i <= 3; i++)
        {
            string baseName = $"DocumentPair{i}";

            // Version A
            var docA = new Document();
            var builderA = new DocumentBuilder(docA);
            builderA.Writeln($"This is {baseName} version A. It contains original content.");
            string fileA = Path.Combine(inputDir, $"{baseName}_A.docx");
            docA.Save(fileA);

            // Version B – slightly different text to generate revisions.
            var docB = new Document();
            var builderB = new DocumentBuilder(docB);
            builderB.Writeln($"This is {baseName} version B. It contains modified content.");
            string fileB = Path.Combine(inputDir, $"{baseName}_B.docx");
            docB.Save(fileB);
        }

        // Process each pair: compare version A with version B, track revisions, and save the result.
        var versionAFiles = Directory.GetFiles(inputDir, "*_A.docx");
        foreach (var fileA in versionAFiles)
        {
            // Derive the corresponding version B file name.
            string fileB = fileA.Replace("_A.docx", "_B.docx");
            if (!File.Exists(fileB))
                continue; // Skip if the matching B file is missing.

            // Load the two documents.
            var docA = new Document(fileA);
            var docB = new Document(fileB);

            // Perform comparison. Revisions will be added to docA.
            docA.Compare(docB, "BatchUser", DateTime.Now);

            // Verify that revisions were created.
            int revisionCount = docA.Revisions.Count;
            if (revisionCount == 0)
                throw new InvalidOperationException($"No revisions detected for pair {Path.GetFileNameWithoutExtension(fileA)}.");

            // Save the compared document to the output folder.
            string resultFileName = Path.GetFileNameWithoutExtension(fileA).Replace("_A", "_Compared") + ".docx";
            string resultPath = Path.Combine(outputDir, resultFileName);
            docA.Save(resultPath);

            // Optional: write a simple console report (no user interaction required).
            Console.WriteLine($"{Path.GetFileName(fileA)} vs {Path.GetFileName(fileB)} -> {revisionCount} revisions saved to {resultFileName}");
        }

        // Indicate completion.
        Console.WriteLine("Batch comparison completed.");
    }
}
