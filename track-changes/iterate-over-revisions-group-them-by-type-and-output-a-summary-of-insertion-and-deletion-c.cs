using System;
using System.Collections.Generic;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial text – this will NOT be a revision.
        builder.Write("Original text. ");

        // Start tracking revisions.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Insert new text – this will be recorded as an insertion revision.
        builder.Write("Inserted text. ");

        // Delete the original run – this will be recorded as a deletion revision.
        Run originalRun = doc.FirstSection.Body.FirstParagraph.Runs[0];
        originalRun.Remove();

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Count revisions by type.
        Dictionary<RevisionType, int> revisionCounts = new Dictionary<RevisionType, int>();
        foreach (Revision rev in doc.Revisions)
        {
            if (!revisionCounts.ContainsKey(rev.RevisionType))
                revisionCounts[rev.RevisionType] = 0;
            revisionCounts[rev.RevisionType]++;
        }

        int insertionCount = revisionCounts.ContainsKey(RevisionType.Insertion) ? revisionCounts[RevisionType.Insertion] : 0;
        int deletionCount = revisionCounts.ContainsKey(RevisionType.Deletion) ? revisionCounts[RevisionType.Deletion] : 0;

        // Output the summary.
        Console.WriteLine($"Insertions: {insertionCount}");
        Console.WriteLine($"Deletions: {deletionCount}");

        // Save the document (optional, demonstrates file output).
        doc.Save("RevisionsSummary.docx");
    }
}
