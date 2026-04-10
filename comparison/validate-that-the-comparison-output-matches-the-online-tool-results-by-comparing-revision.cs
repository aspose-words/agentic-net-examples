using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document with deterministic content.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This is the original paragraph.");

        // Create the edited document with a clear difference.
        Document edited = new Document();
        DocumentBuilder builderEdited = new DocumentBuilder(edited);
        builderEdited.Writeln("This is the edited paragraph with extra text.");

        // Ensure both documents have no revisions before comparison, then compare.
        if (original.Revisions.Count == 0 && edited.Revisions.Count == 0)
        {
            original.Compare(edited, "Comparer", DateTime.Now);
        }

        // Save the comparison result (original now contains revisions).
        const string outputPath = "ComparisonResult.docx";
        original.Save(outputPath);

        // Inspect revisions.
        int revisionCount = original.Revisions.Count;
        int insertionCount = 0;
        int deletionCount = 0;

        foreach (Revision rev in original.Revisions)
        {
            if (rev.RevisionType == RevisionType.Insertion)
                insertionCount++;
            else if (rev.RevisionType == RevisionType.Deletion)
                deletionCount++;
        }

        // Validate that revisions exist and both insertion and deletion types are present.
        bool validationPassed = revisionCount > 0 && insertionCount > 0 && deletionCount > 0;

        // Output the validation results.
        Console.WriteLine($"Revisions count: {revisionCount}");
        Console.WriteLine($"Insertions: {insertionCount}, Deletions: {deletionCount}");
        Console.WriteLine($"Validation result: {(validationPassed ? "PASS" : "FAIL")}");
    }
}
