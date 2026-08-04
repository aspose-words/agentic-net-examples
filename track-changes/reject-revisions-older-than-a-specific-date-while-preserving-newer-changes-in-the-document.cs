using System;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define the cutoff date. Revisions older than this will be rejected.
        DateTime cutoffDate = new DateTime(2023, 1, 1);

        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add initial content that is not a revision.
        builder.Writeln("Initial content.");

        // Create an old revision (date before the cutoff).
        doc.StartTrackRevisions("Alice", new DateTime(2022, 12, 15));
        builder.Writeln("This is an old revision.");
        doc.StopTrackRevisions();

        // Create a new revision (date after the cutoff).
        doc.StartTrackRevisions("Bob", new DateTime(2023, 2, 10));
        builder.Writeln("This is a new revision.");
        doc.StopTrackRevisions();

        // Reject revisions older than the cutoff date.
        foreach (Revision rev in doc.Revisions.ToList())
        {
            if (rev.DateTime < cutoffDate)
                rev.Reject();
        }

        // Save the document with only the newer revisions preserved.
        doc.Save("RevisionsFiltered.docx");
    }
}
