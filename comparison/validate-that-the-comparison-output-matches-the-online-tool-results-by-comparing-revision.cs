using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document with deterministic content.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world.");

        // Create the revised document with a clear difference that produces both insertion and deletion revisions.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        // Replace "Hello" with "Goodbye" to generate a deletion (Hello) and an insertion (Goodbye).
        builderRevised.Writeln("Goodbye world.");

        // Perform the comparison. The original document will receive revisions.
        original.Compare(revised, "Tester", DateTime.Now);

        // Inspect the revisions collection.
        int revisionCount = original.Revisions.Count;
        bool hasInsertion = original.Revisions.Any(r => r.RevisionType == RevisionType.Insertion);
        bool hasDeletion = original.Revisions.Any(r => r.RevisionType == RevisionType.Deletion);

        // Validate that the comparison produced the expected revisions.
        if (revisionCount < 2 || !hasInsertion || !hasDeletion)
        {
            throw new InvalidOperationException(
                $"Unexpected revision results. Count: {revisionCount}, " +
                $"HasInsertion: {hasInsertion}, HasDeletion: {hasDeletion}");
        }

        // Save the compared document for visual verification.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonResult.docx");
        original.Save(outputPath);

        // Write a simple text report summarizing the revisions.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "RevisionReport.txt");
        File.WriteAllText(reportPath,
            $"Revision count: {revisionCount}{Environment.NewLine}" +
            $"Contains Insertion: {hasInsertion}{Environment.NewLine}" +
            $"Contains Deletion: {hasDeletion}{Environment.NewLine}" +
            $"Output document: {outputPath}");
    }
}
