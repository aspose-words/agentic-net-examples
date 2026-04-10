using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Base directory for all generated files.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");

        // Ensure folders exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Number of document pairs to generate.
        const int pairCount = 2;

        // Create sample document pairs.
        for (int i = 1; i <= pairCount; i++)
        {
            // Paths for the original and edited documents.
            string originalPath = Path.Combine(inputDir, $"Original_{i}.docx");
            string editedPath = Path.Combine(inputDir, $"Edited_{i}.docx");

            // Create the original document with deterministic content.
            Document original = new Document();
            DocumentBuilder builder = new DocumentBuilder(original);
            builder.Writeln($"This is the original content for pair {i}.");
            original.Save(originalPath);

            // Clone the original, modify it to introduce differences, and save as edited.
            Document edited = (Document)original.Clone(true);
            DocumentBuilder editedBuilder = new DocumentBuilder(edited);
            editedBuilder.Writeln($"Additional line for pair {i} to create a revision.");
            edited.Save(editedPath);
        }

        // Batch process each pair: compare and save the result with revisions.
        for (int i = 1; i <= pairCount; i++)
        {
            string originalPath = Path.Combine(inputDir, $"Original_{i}.docx");
            string editedPath = Path.Combine(inputDir, $"Edited_{i}.docx");

            // Load documents.
            Document originalDoc = new Document(originalPath);
            Document editedDoc = new Document(editedPath);

            // Ensure both documents have no pre‑existing revisions before comparison.
            if (originalDoc.Revisions.Count == 0 && editedDoc.Revisions.Count == 0)
            {
                // Perform comparison; revisions will be added to the original document.
                originalDoc.Compare(editedDoc, "BatchUser", DateTime.Now);
            }

            // Verify that revisions were generated.
            int revisionCount = originalDoc.Revisions.Count;
            Console.WriteLine($"Pair {i}: {revisionCount} revision(s) detected.");

            // Save the compared document (which now contains revision tracking).
            string comparedPath = Path.Combine(outputDir, $"Compared_{i}.docx");
            originalDoc.Save(comparedPath);
        }
    }
}
