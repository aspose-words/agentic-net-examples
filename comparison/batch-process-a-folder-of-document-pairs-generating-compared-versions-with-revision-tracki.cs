using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class BatchDocumentComparison
{
    public static void Main()
    {
        // Define the input folder relative to the current directory.
        string inputDir = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonInput");
        Directory.CreateDirectory(inputDir);

        // Create sample document pairs with deterministic content.
        CreateSampleDocument(Path.Combine(inputDir, "docA_v1.docx"), "Document A - Version 1");
        CreateSampleDocument(Path.Combine(inputDir, "docA_v2.docx"), "Document A - Version 2 (modified)");
        CreateSampleDocument(Path.Combine(inputDir, "docB_v1.docx"), "Document B - Initial content");
        CreateSampleDocument(Path.Combine(inputDir, "docB_v2.docx"), "Document B - Updated content with changes");

        // Process each pair: compare version 1 with version 2 and save the result.
        string[] baseNames = { "docA", "docB" };
        foreach (string baseName in baseNames)
        {
            string firstPath = Path.Combine(inputDir, $"{baseName}_v1.docx");
            string secondPath = Path.Combine(inputDir, $"{baseName}_v2.docx");

            // Load the two documents.
            Document firstDoc = new Document(firstPath);
            Document secondDoc = new Document(secondPath);

            // Perform comparison with revision tracking.
            firstDoc.Compare(secondDoc, "BatchUser", DateTime.Now);

            // Verify that revisions were generated.
            if (firstDoc.Revisions.Count == 0)
                throw new InvalidOperationException($"Expected revisions for {baseName}, but none were found.");

            // Save the compared document.
            string resultPath = Path.Combine(inputDir, $"{baseName}_compared.docx");
            firstDoc.Save(resultPath);
        }

        // Optional: output a simple summary.
        Console.WriteLine($"Batch comparison completed. Results are stored in: {inputDir}");
    }

    // Helper method to create a single‑page document with the specified text.
    private static void CreateSampleDocument(string filePath, string content)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(content);
        doc.Save(filePath);
    }
}
