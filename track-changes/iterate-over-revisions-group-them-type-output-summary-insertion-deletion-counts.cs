using System;
using System.IO;
using System.Linq;
using Aspose.Words;

class RevisionSummary
{
    static void Main()
    {
        // Load the document that contains tracked changes.
        // (Assumes the file "Revisions.docx" exists in the working directory.)
        Document doc = new Document("Revisions.docx");

        // Group all revisions by their RevisionType.
        var revisionGroups = doc.Revisions
                                .GroupBy(r => r.RevisionType)
                                .ToDictionary(g => g.Key, g => g.Count());

        // Retrieve counts for insertions and deletions, defaulting to 0 if the type is absent.
        int insertionCount = revisionGroups.ContainsKey(RevisionType.Insertion)
                             ? revisionGroups[RevisionType.Insertion] : 0;
        int deletionCount = revisionGroups.ContainsKey(RevisionType.Deletion)
                            ? revisionGroups[RevisionType.Deletion] : 0;

        // Output the summary to the console.
        Console.WriteLine($"Insertions: {insertionCount}");
        Console.WriteLine($"Deletions: {deletionCount}");

        // Optionally, write the summary to a text file.
        string summaryPath = "RevisionSummary.txt";
        File.WriteAllText(summaryPath,
            $"Insertions: {insertionCount}{Environment.NewLine}Deletions: {deletionCount}");
    }
}
