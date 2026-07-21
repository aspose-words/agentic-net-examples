using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class BatchDocumentComparison
{
    public static void Main()
    {
        // Define the folder that will hold the sample documents and comparison results.
        string inputDir = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonInput");
        Directory.CreateDirectory(inputDir);

        // Create a few pairs of documents with deterministic differences.
        for (int i = 1; i <= 3; i++)
        {
            // Original version.
            Document original = new Document();
            DocumentBuilder builderOrig = new DocumentBuilder(original);
            builderOrig.Writeln($"Document {i} - Original version.");
            string originalPath = Path.Combine(inputDir, $"Doc{i}_v1.docx");
            original.Save(originalPath);

            // Revised version with an extra line to ensure a revision is generated.
            Document revised = new Document();
            DocumentBuilder builderRev = new DocumentBuilder(revised);
            builderRev.Writeln($"Document {i} - Original version.");
            builderRev.Writeln($"Document {i} - Revised additional line.");
            string revisedPath = Path.Combine(inputDir, $"Doc{i}_v2.docx");
            revised.Save(revisedPath);
        }

        // Process each pair in the folder.
        string[] originalFiles = Directory.GetFiles(inputDir, "*_v1.docx");
        foreach (string originalFile in originalFiles)
        {
            string baseName = Path.GetFileNameWithoutExtension(originalFile);
            string revisedFile = Path.Combine(inputDir, baseName.Replace("_v1", "_v2") + ".docx");

            if (!File.Exists(revisedFile))
                continue; // Skip if the counterpart does not exist.

            // Load the documents.
            Document docOriginal = new Document(originalFile);
            Document docRevised = new Document(revisedFile);

            // Perform the comparison, tracking revisions.
            docOriginal.Compare(docRevised, "BatchUser", DateTime.Now);

            // Verify that at least one revision was created.
            if (docOriginal.Revisions.Count == 0)
                throw new InvalidOperationException($"No revisions detected for pair {baseName}.");

            // Save the compared document with revisions.
            string resultPath = Path.Combine(inputDir, $"{baseName}_compared.docx");
            docOriginal.Save(resultPath);

            Console.WriteLine($"Compared {Path.GetFileName(originalFile)} with {Path.GetFileName(revisedFile)} -> {Path.GetFileName(resultPath)} (Revisions: {docOriginal.Revisions.Count})");
        }

        Console.WriteLine("Batch comparison completed.");
    }
}
