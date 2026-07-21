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

        // Add some initial text that will not be tracked as a revision.
        builder.Write("This will be deleted. ");
        builder.Write("This stays unchanged. ");

        // Start tracking revisions.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Insert new text – this will be recorded as an insertion revision.
        builder.Write("Inserted revision text. ");

        // Delete the first run (the text added before tracking) – this will be a deletion revision.
        // The first run is at index 0 of the paragraph's Runs collection.
        doc.FirstSection.Body.FirstParagraph.Runs[0].Remove();

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Save the document (optional, but complies with the rule to save when output is expected).
        doc.Save("RevisionsSummary.docx");

        // Count revisions by type.
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
