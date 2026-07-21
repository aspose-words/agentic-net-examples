using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class ComparisonValidation
{
    public static void Main()
    {
        // Create the original document with a single paragraph.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Alpha");

        // Create the revised document with different content.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Beta");

        // Perform the comparison. The original document will receive revisions.
        original.Compare(revised, "Tester", DateTime.Now);

        // Expected: one deletion (Alpha) and one insertion (Beta) => 2 revisions.
        int expectedCount = 2;
        if (original.Revisions.Count != expectedCount)
            throw new InvalidOperationException($"Expected {expectedCount} revisions, but found {original.Revisions.Count}.");

        bool hasDeletion = false;
        bool hasInsertion = false;

        foreach (Revision rev in original.Revisions)
        {
            if (rev.RevisionType == RevisionType.Deletion)
                hasDeletion = true;
            else if (rev.RevisionType == RevisionType.Insertion)
                hasInsertion = true;
        }

        if (!hasDeletion || !hasInsertion)
            throw new InvalidOperationException("Revisions do not contain both deletion and insertion types as expected.");

        // Save the document that now contains the revisions.
        string outputPath = "ComparisonResult.docx";
        original.Save(outputPath);

        // Report the validation result.
        Console.WriteLine($"Comparison successful. Revision count: {original.Revisions.Count}");
        Console.WriteLine($"Contains Deletion: {hasDeletion}, Contains Insertion: {hasInsertion}");
        Console.WriteLine($"Result saved to: {outputPath}");
    }
}
