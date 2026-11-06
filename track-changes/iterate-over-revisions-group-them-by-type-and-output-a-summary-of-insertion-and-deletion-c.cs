using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add initial content that will later be deleted.
        builder.Writeln("Original paragraph that will be deleted.");

        // Enable track changes.
        doc.StartTrackRevisions("Demo Author", DateTime.Now);

        // Insert a new paragraph – this creates an insertion revision.
        builder.Writeln("Inserted paragraph.");

        // Delete the original paragraph – this creates a deletion revision.
        Paragraph paragraphToDelete = doc.FirstSection.Body.Paragraphs[0];
        paragraphToDelete.Remove();

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Save the document (demonstrates file output).
        doc.Save("RevisionsDemo.docx");

        // Iterate over revisions and count insertions and deletions.
        int insertionCount = 0;
        int deletionCount = 0;

        foreach (Revision rev in doc.Revisions)
        {
            if (rev.RevisionType == RevisionType.Insertion)
                insertionCount++;
            else if (rev.RevisionType == RevisionType.Deletion)
                deletionCount++;
        }

        // Output the summary.
        Console.WriteLine($"Insertions: {insertionCount}");
        Console.WriteLine($"Deletions: {deletionCount}");
    }
}
