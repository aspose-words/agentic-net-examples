using System;
using Aspose.Words;

class RevisionDemo
{
    static void Main()
    {
        // Load an existing document from disk.
        Document doc = new Document("Input.docx");

        // Enable tracking of changes. All subsequent edits will be recorded as revisions.
        doc.StartTrackRevisions("DemoUser");

        // Perform a find-and-replace operation on the whole document.
        // This will create a revision for each replacement made.
        int replacements = doc.Range.Replace("_FullName_", "John Doe");

        // Stop tracking so further edits are not recorded.
        doc.StopTrackRevisions();

        // Save the modified document (optional, shows the result with revisions).
        doc.Save("OutputWithRevisions.docx");

        // List all generated revisions.
        Console.WriteLine($"Total replacements made: {replacements}");
        Console.WriteLine($"Total revisions created: {doc.Revisions.Count}");

        int revNumber = 0;
        foreach (Revision rev in doc.Revisions)
        {
            revNumber++;
            Console.WriteLine($"Revision #{revNumber}:");
            Console.WriteLine($"  Author   : {rev.Author}");
            Console.WriteLine($"  Type     : {rev.RevisionType}");
            Console.WriteLine($"  DateTime : {rev.DateTime}");

            // Some revisions (e.g., deletions) may not have a parent node that contains text.
            string affectedText = rev.ParentNode != null ? rev.ParentNode.GetText().Trim() : "<no text>";
            Console.WriteLine($"  Affected text: \"{affectedText}\"");
        }
    }
}
