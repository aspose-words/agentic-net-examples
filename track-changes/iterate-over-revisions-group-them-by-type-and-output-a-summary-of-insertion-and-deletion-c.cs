using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add an initial paragraph that will later be deleted.
        builder.Writeln("This paragraph will be deleted.");

        // Enable track changes.
        doc.StartTrackRevisions("Sample Author", DateTime.Now);

        // Insert a new paragraph – this will be recorded as an insertion revision.
        builder.Writeln("This paragraph was inserted.");

        // Delete the original paragraph – this will be recorded as a deletion revision.
        Node firstParagraph = doc.FirstSection.Body.FirstParagraph;
        if (firstParagraph != null)
        {
            firstParagraph.Remove();
        }

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Save the document (demonstrates that revisions are persisted).
        doc.Save("RevisionsDemo.docx");

        // Iterate over revisions and count insertions and deletions.
        int insertionCount = 0;
        int deletionCount = 0;

        foreach (Revision rev in doc.Revisions)
        {
            switch (rev.RevisionType)
            {
                case RevisionType.Insertion:
                    insertionCount++;
                    break;
                case RevisionType.Deletion:
                    deletionCount++;
                    break;
                // Other revision types are ignored for this summary.
            }
        }

        // Output the summary.
        Console.WriteLine($"Insertions: {insertionCount}");
        Console.WriteLine($"Deletions: {deletionCount}");
    }
}
