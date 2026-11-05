using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define a folder for input and output documents.
        string inputDir = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonInput");
        Directory.CreateDirectory(inputDir);

        // Create sample document pairs with deterministic differences.
        for (int i = 1; i <= 2; i++)
        {
            // First version of the pair.
            Document docV1 = new();
            DocumentBuilder builderV1 = new(docV1);
            builderV1.Writeln($"Document {i} - Version 1");
            builderV1.Writeln($"This is the original content for document pair {i}.");
            string v1Path = Path.Combine(inputDir, $"doc{i}_v1.docx");
            docV1.Save(v1Path);

            // Second version of the pair with a slight change.
            Document docV2 = new();
            DocumentBuilder builderV2 = new(docV2);
            builderV2.Writeln($"Document {i} - Version 2");
            builderV2.Writeln($"This is the modified content for document pair {i}.");
            string v2Path = Path.Combine(inputDir, $"doc{i}_v2.docx");
            docV2.Save(v2Path);
        }

        // Process each document pair: compare and save the result with revisions.
        for (int i = 1; i <= 2; i++)
        {
            string v1Path = Path.Combine(inputDir, $"doc{i}_v1.docx");
            string v2Path = Path.Combine(inputDir, $"doc{i}_v2.docx");

            // Load the two versions.
            Document baseDoc = new(v1Path);
            Document revisedDoc = new(v2Path);

            // Perform comparison. Revisions will be tracked in baseDoc.
            baseDoc.Compare(revisedDoc, "BatchComparer", DateTime.Now);

            // Verify that revisions were created.
            int revisionCount = baseDoc.Revisions.Count;
            if (revisionCount == 0)
            {
                throw new InvalidOperationException($"Expected revisions for document pair {i}, but none were found.");
            }

            // Save the compared document.
            string resultPath = Path.Combine(inputDir, $"doc{i}_comparison.docx");
            baseDoc.Save(resultPath);
        }
    }
}
