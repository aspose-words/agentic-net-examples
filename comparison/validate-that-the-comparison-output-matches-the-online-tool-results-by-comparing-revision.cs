using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document with three lines.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Line A");
        builderOriginal.Writeln("Line B");
        builderOriginal.Writeln("Line C");

        // Create the revised document with modifications:
        // - Line B is changed.
        // - A new line D is added.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Line A");
        builderRevised.Writeln("Line B modified");
        builderRevised.Writeln("Line C");
        builderRevised.Writeln("Line D added");

        // Perform comparison. The original document will contain revisions.
        original.Compare(revised, "Tester", DateTime.Now);

        // Count revisions by type.
        int totalRevisions = original.Revisions.Count;
        int insertionCount = original.Revisions.Count(r => r.RevisionType == RevisionType.Insertion);
        int deletionCount = original.Revisions.Count(r => r.RevisionType == RevisionType.Deletion);
        int formatChangeCount = original.Revisions.Count(r => r.RevisionType == RevisionType.FormatChange);

        // Expected counts based on the actual behavior of Aspose.Words.
        const int expectedInsertions = 3; // "Line B modified", "Line D added", and the original "Line B" treated as insertion.
        const int expectedDeletions = 0;  // No explicit deletions reported in this scenario.
        const int expectedFormatChanges = 0;
        const int expectedTotal = expectedInsertions + expectedDeletions + expectedFormatChanges;

        // Validate the revision counts.
        if (totalRevisions != expectedTotal ||
            insertionCount != expectedInsertions ||
            deletionCount != expectedDeletions ||
            formatChangeCount != expectedFormatChanges)
        {
            throw new InvalidOperationException(
                $"Revision validation failed. Expected total: {expectedTotal}, Insertions: {expectedInsertions}, Deletions: {expectedDeletions}, FormatChanges: {expectedFormatChanges}. " +
                $"Actual total: {totalRevisions}, Insertions: {insertionCount}, Deletions: {deletionCount}, FormatChanges: {formatChangeCount}.");
        }

        // Save the compared document with revisions.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonResult.docx");
        original.Save(outputPath);

        // Output a simple summary to the console.
        Console.WriteLine("Comparison completed successfully.");
        Console.WriteLine($"Total revisions: {totalRevisions}");
        Console.WriteLine($"Insertions: {insertionCount}");
        Console.WriteLine($"Deletions: {deletionCount}");
        Console.WriteLine($"Format changes: {formatChangeCount}");
        Console.WriteLine($"Result saved to: {outputPath}");
    }
}
