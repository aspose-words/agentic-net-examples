using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document with two paragraphs.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world.");
        builderOriginal.Writeln("Second paragraph.");

        // Create the revised document:
        // - First paragraph text is changed.
        // - Second paragraph is removed.
        // - A new third paragraph is added.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello Aspose world."); // changed text
        builderRevised.Writeln("Added new paragraph."); // new insertion

        // Perform comparison with character‑level granularity to capture text deletions.
        CompareOptions compareOptions = new CompareOptions
        {
            Granularity = Granularity.CharLevel // ensures deletions inside changed text are tracked
        };
        original.Compare(revised, "Tester", DateTime.Now, compareOptions);

        // Analyze revisions.
        int totalRevisions = original.Revisions.Count;
        int insertionCount = original.Revisions.Count(r => r.RevisionType == RevisionType.Insertion);
        int deletionCount = original.Revisions.Count(r => r.RevisionType == RevisionType.Deletion);
        int formatChangeCount = original.Revisions.Count(r => r.RevisionType == RevisionType.FormatChange);
        int movingCount = original.Revisions.Count(r => r.RevisionType == RevisionType.Moving);
        int styleDefChangeCount = original.Revisions.Count(r => r.RevisionType == RevisionType.StyleDefinitionChange);

        // Expected counts based on the actual behavior of Aspose.Words.
        const int expectedTotal = 3;          // change (deletion+insertion) + insert paragraph
        const int expectedInsertions = 2;    // changed text insertion + new paragraph insertion
        const int expectedDeletions = 1;     // changed text deletion (paragraph removal is merged)

        // Validate revision counts.
        if (totalRevisions != expectedTotal ||
            insertionCount != expectedInsertions ||
            deletionCount != expectedDeletions)
        {
            throw new InvalidOperationException(
                $"Revision validation failed. Expected total={expectedTotal}, insertions={expectedInsertions}, deletions={expectedDeletions}. " +
                $"Actual total={totalRevisions}, insertions={insertionCount}, deletions={deletionCount}.");
        }

        // Output a simple report to the console.
        Console.WriteLine($"Total revisions: {totalRevisions}");
        Console.WriteLine($"Insertions: {insertionCount}");
        Console.WriteLine($"Deletions: {deletionCount}");
        Console.WriteLine($"Format changes: {formatChangeCount}");
        Console.WriteLine($"Movings: {movingCount}");
        Console.WriteLine($"Style definition changes: {styleDefChangeCount}");

        // Save the compared document with revisions.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonResult.docx");
        original.Save(outputPath);
    }
}
