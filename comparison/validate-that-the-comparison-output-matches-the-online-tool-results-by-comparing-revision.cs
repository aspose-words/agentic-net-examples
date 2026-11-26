using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document with three paragraphs.
        Document original = new Document();
        DocumentBuilder builderOrig = new DocumentBuilder(original);
        builderOrig.Writeln("Paragraph 1.");
        builderOrig.Writeln("Paragraph 2.");
        builderOrig.Writeln("Paragraph 3.");

        // Create the revised document with differences:
        // - Insert a new paragraph.
        // - Modify the text of the second paragraph.
        // - Delete the third paragraph.
        Document revised = new Document();
        DocumentBuilder builderRev = new DocumentBuilder(revised);
        builderRev.Writeln("Paragraph 1.");
        builderRev.Writeln("Inserted paragraph."); // New insertion.
        builderRev.Writeln("Paragraph 2 modified."); // Text change (deletion + insertion).

        // Perform comparison.
        original.Compare(revised, "TestAuthor", DateTime.Now);

        // Count revisions by type.
        int insertionCount = 0;
        int deletionCount = 0;
        int formatChangeCount = 0;

        foreach (Revision rev in original.Revisions)
        {
            switch (rev.RevisionType)
            {
                case RevisionType.Insertion:
                    insertionCount++;
                    break;
                case RevisionType.Deletion:
                    deletionCount++;
                    break;
                case RevisionType.FormatChange:
                    formatChangeCount++;
                    break;
            }
        }

        // Expected revision counts based on the known differences.
        const int expectedInsertions = 2; // Inserted paragraph + new text for modified paragraph.
        const int expectedDeletions = 2;  // Deleted paragraph + old text of modified paragraph.
        const int expectedFormatChanges = 0;

        // Validate that the actual counts match the expected counts.
        if (insertionCount != expectedInsertions ||
            deletionCount != expectedDeletions ||
            formatChangeCount != expectedFormatChanges)
        {
            throw new InvalidOperationException(
                $"Revision counts mismatch. " +
                $"Insertions: {insertionCount}/{expectedInsertions}, " +
                $"Deletions: {deletionCount}/{expectedDeletions}, " +
                $"FormatChanges: {formatChangeCount}/{expectedFormatChanges}");
        }

        // Save the comparison result.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonResult.docx");
        original.Save(outputPath);

        // Report the outcome.
        Console.WriteLine("Comparison completed successfully.");
        Console.WriteLine($"Insertions: {insertionCount}");
        Console.WriteLine($"Deletions: {deletionCount}");
        Console.WriteLine($"Format changes: {formatChangeCount}");
        Console.WriteLine($"Result saved to: {outputPath}");
    }
}
