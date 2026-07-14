using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Set up a deterministic folder for the example.
        string inputDir = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonInput");
        Directory.CreateDirectory(inputDir);

        // Create sample document pairs.
        CreateSamplePair(inputDir, "pair1", "Original content for pair 1.", "Modified content for pair 1.");
        CreateSamplePair(inputDir, "pair2", "First version of pair 2.", "Second version of pair 2 with change.");
        CreateSamplePair(inputDir, "pair3", "Alpha version.", "Beta version with differences.");

        // Process each pair: compare and save the result with revisions.
        foreach (var pair in GetDocumentPairs(inputDir))
        {
            // Load the original and the modified document.
            Document original = new Document(pair.OriginalPath);
            Document revised = new Document(pair.ModifiedPath);

            // Perform comparison. Provide author and timestamp.
            original.Compare(revised, "BatchProcessor", DateTime.Now);

            // Verify that revisions were generated.
            if (original.Revisions.Count == 0)
                throw new InvalidOperationException($"No revisions detected for pair '{pair.Prefix}'.");

            // Save the compared document.
            string resultPath = Path.Combine(inputDir, $"{pair.Prefix}_compared.docx");
            original.Save(resultPath);
        }
    }

    // Creates a pair of documents with given texts.
    private static void CreateSamplePair(string folder, string prefix, string originalText, string modifiedText)
    {
        string originalPath = Path.Combine(folder, $"{prefix}_original.docx");
        string modifiedPath = Path.Combine(folder, $"{prefix}_modified.docx");

        CreateDocument(originalPath, originalText);
        CreateDocument(modifiedPath, modifiedText);
    }

    // Helper to create a single document with specified text.
    private static void CreateDocument(string path, string text)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(text);
        doc.Save(path);
    }

    // Represents a pair of document file paths.
    private class DocumentPair
    {
        public string Prefix { get; set; } = "";
        public string OriginalPath { get; set; } = "";
        public string ModifiedPath { get; set; } = "";
    }

    // Retrieves all document pairs from the folder based on naming convention.
    private static IEnumerable<DocumentPair> GetDocumentPairs(string folder)
    {
        var files = Directory.GetFiles(folder, "*.docx");
        var pairs = new Dictionary<string, DocumentPair>(StringComparer.OrdinalIgnoreCase);

        foreach (var file in files)
        {
            string fileName = Path.GetFileNameWithoutExtension(file);
            // Expected format: {prefix}_original or {prefix}_modified
            int underscoreIndex = fileName.LastIndexOf('_');
            if (underscoreIndex <= 0) continue;

            string prefix = fileName.Substring(0, underscoreIndex);
            string suffix = fileName.Substring(underscoreIndex + 1).ToLowerInvariant();

            if (!pairs.TryGetValue(prefix, out var pair))
            {
                pair = new DocumentPair { Prefix = prefix };
                pairs[prefix] = pair;
            }

            if (suffix == "original")
                pair.OriginalPath = file;
            else if (suffix == "modified")
                pair.ModifiedPath = file;
        }

        foreach (var pair in pairs.Values)
        {
            // Ensure both files exist before yielding.
            if (!string.IsNullOrEmpty(pair.OriginalPath) && !string.IsNullOrEmpty(pair.ModifiedPath))
                yield return pair;
        }
    }
}
