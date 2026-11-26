using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class BatchComparisonExample
{
    public static void Main()
    {
        // Prepare input and output directories.
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
            // First version of the pair.
            Document docA = new Document();
            DocumentBuilder builderA = new DocumentBuilder(docA);
            builderA.Writeln($"Document {i} - Version A");
            builderA.Writeln($"Common line for pair {i}.");
            string pathA = Path.Combine(inputDir, $"Doc{i}_A.docx");
            docA.Save(pathA);

            // Second version of the pair with a slight change.
            Document docB = new Document();
            DocumentBuilder builderB = new DocumentBuilder(docB);
            builderB.Writeln($"Document {i} - Version B"); // Different heading.
            builderB.Writeln($"Common line for pair {i}.");
            string pathB = Path.Combine(inputDir, $"Doc{i}_B.docx");
            docB.Save(pathB);
        }

        // Process each pair: compare and save the result with revisions.
        for (int i = 1; i <= pairCount; i++)
        {
            string pathA = Path.Combine(inputDir, $"Doc{i}_A.docx");
            string pathB = Path.Combine(inputDir, $"Doc{i}_B.docx");

            // Load the documents.
            Document docA = new Document(pathA);
            Document docB = new Document(pathB);

            // Perform comparison. Revisions will be stored in docA.
            docA.Compare(docB, "BatchUser", DateTime.Now);

            // Verify that revisions were created.
            int revisionCount = docA.Revisions.Count;
            if (revisionCount == 0)
            {
                throw new InvalidOperationException($"Expected revisions for pair {i}, but none were found.");
            }

            // Save the compared document.
            string resultPath = Path.Combine(outputDir, $"Doc{i}_Compared.docx");
            docA.Save(resultPath);
        }

        // Optional: write a simple summary to the console.
        Console.WriteLine($"Batch comparison completed. Processed {pairCount} document pairs.");
        Console.WriteLine($"Input files are located in: {inputDir}");
        Console.WriteLine($"Compared results are saved in: {outputDir}");
    }
}
