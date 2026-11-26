using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document with mixed‑case text.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello World");

        // Create the revised document with the same text but different case.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("hello world");

        // Configure comparison options to ignore case differences.
        // The correct property name is IgnoreCaseChanges.
        CompareOptions options = new CompareOptions
        {
            IgnoreCaseChanges = true
        };

        // Perform the comparison.
        original.Compare(revised, "Comparer", DateTime.Now, options);

        // Verify that no revisions were generated because case differences are ignored.
        int revisionCount = original.Revisions.Count;
        Console.WriteLine($"Number of revisions after case‑insensitive comparison: {revisionCount}");

        // Save the result document.
        string outputPath = "ComparisonResult.docx";
        original.Save(outputPath);
    }
}
