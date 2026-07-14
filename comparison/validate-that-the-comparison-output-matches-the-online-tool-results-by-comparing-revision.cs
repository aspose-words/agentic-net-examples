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
        builderOriginal.Writeln("Hello world.");               // Paragraph 1
        builderOriginal.Writeln("This is the second paragraph."); // Paragraph 2

        // Create the revised document with modifications:
        // - Change text in the first paragraph.
        // - Remove the second paragraph.
        // - Add a new third paragraph.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello brave new world."); // Modified first paragraph
        builderRevised.Writeln("Added new paragraph.");   // New third paragraph

        // Perform comparison. The original document will receive revisions.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Save the comparison result for visual inspection if needed.
        string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonResult.docx");
        original.Save(resultPath);

        // Validate revision count and types.
        int revisionCount = original.Revisions.Count;
        if (revisionCount == 0)
            throw new InvalidOperationException("Expected at least one revision after comparison.");

        // Determine which revision types are present.
        var revisionTypes = original.Revisions
                                    .Select(r => r.RevisionType)
                                    .Distinct()
                                    .ToList();

        // Expect at least one insertion and one deletion.
        bool hasInsertion = revisionTypes.Contains(RevisionType.Insertion);
        bool hasDeletion = revisionTypes.Contains(RevisionType.Deletion);

        if (!hasInsertion || !hasDeletion)
            throw new InvalidOperationException("Expected both insertion and deletion revisions.");

        // Output summary to console.
        Console.WriteLine($"Total revisions: {revisionCount}");
        Console.WriteLine("Revision types present:");
        foreach (RevisionType type in revisionTypes)
        {
            Console.WriteLine($"- {type}");
        }

        // Optional: accept all revisions to transform the original into the revised version.
        original.AcceptAllRevisions();
        string acceptedPath = Path.Combine(Directory.GetCurrentDirectory(), "Accepted.docx");
        original.Save(acceptedPath);
    }
}
