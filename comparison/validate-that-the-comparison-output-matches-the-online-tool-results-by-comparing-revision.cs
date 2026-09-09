using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class ComparisonValidator
{
    public static void Main()
    {
        // Create the original document with deterministic content.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world.");

        // Create the revised document with a clear difference (insertion of extra words).
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello brave new world.");

        // Perform the comparison. The original document will receive revisions.
        original.Compare(revised, "Validator", DateTime.Now);

        // Validate that at least one revision was created.
        int revisionCount = original.Revisions?.Count ?? 0;
        if (revisionCount == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison, but none were found.");
        }

        // For this simple text change we expect a single insertion revision.
        if (revisionCount != 1)
        {
            throw new InvalidOperationException($"Expected exactly 1 revision, but found {revisionCount}.");
        }

        // Verify the type of the revision.
        Revision revision = original.Revisions[0];
        if (revision.RevisionType != RevisionType.Insertion)
        {
            throw new InvalidOperationException($"Expected revision type Insertion, but found {revision.RevisionType}.");
        }

        // Save the compared document so the revisions can be inspected manually if needed.
        string outputPath = "ComparisonResult.docx";
        original.Save(outputPath);

        // Inform the user (via console) that validation succeeded.
        Console.WriteLine($"Comparison validation succeeded. Revisions count: {revisionCount}, type: {revision.RevisionType}");
        Console.WriteLine($"Compared document saved to: {outputPath}");
    }
}
