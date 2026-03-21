using System;
using System.IO;
using System.Linq;
using Aspose.Words;

class RevisionSummary
{
    static void Main()
    {
        // Create a new document and add some tracked changes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start tracking revisions.
        doc.StartTrackRevisions("Author");

        // Insertion revision: add a paragraph.
        builder.Writeln("This is an inserted paragraph.");

        // Add another paragraph that we will later delete.
        builder.Writeln("This paragraph will be deleted.");

        // Stop tracking to finalize the insertion.
        doc.StopTrackRevisions();

        // Delete the second paragraph while tracking revisions.
        // Re‑enable tracking for the deletion.
        doc.StartTrackRevisions("Author");
        Paragraph paraToDelete = doc.FirstSection.Body.Paragraphs[2]; // zero‑based index
        paraToDelete.Remove(); // This creates a deletion revision.
        doc.StopTrackRevisions();

        // Group all revisions by their RevisionType.
        var revisionGroups = doc.Revisions
                                .GroupBy(r => r.RevisionType)
                                .ToDictionary(g => g.Key, g => g.Count());

        // Retrieve counts for insertions and deletions, defaulting to 0 if the type is absent.
        int insertionCount = revisionGroups.TryGetValue(RevisionType.Insertion, out var ins) ? ins : 0;
        int deletionCount = revisionGroups.TryGetValue(RevisionType.Deletion, out var del) ? del : 0;

        // Output the summary to the console.
        Console.WriteLine($"Insertions: {insertionCount}");
        Console.WriteLine($"Deletions: {deletionCount}");

        // Optionally, write the summary to a text file.
        string summaryPath = "RevisionSummary.txt";
        File.WriteAllText(summaryPath,
            $"Insertions: {insertionCount}{Environment.NewLine}Deletions: {deletionCount}");
    }
}
