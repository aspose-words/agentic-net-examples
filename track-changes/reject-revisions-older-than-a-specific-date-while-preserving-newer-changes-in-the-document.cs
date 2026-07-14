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

        // Add some initial content that is not a revision.
        builder.Writeln("Original paragraph. ");

        // Track an older revision (10 days ago).
        DateTime oldRevisionDate = DateTime.Now.AddDays(-10);
        doc.StartTrackRevisions("Author1", oldRevisionDate);
        builder.Writeln("This text is an old revision. ");
        doc.StopTrackRevisions();

        // Track a newer revision (1 day ago).
        DateTime newRevisionDate = DateTime.Now.AddDays(-1);
        doc.StartTrackRevisions("Author2", newRevisionDate);
        builder.Writeln("This text is a new revision. ");
        doc.StopTrackRevisions();

        // Define the cutoff date: revisions older than this will be rejected.
        DateTime cutoffDate = DateTime.Now.AddDays(-5);

        // Collect revisions that are older than the cutoff.
        List<Revision> revisionsToReject = new List<Revision>();
        foreach (Revision rev in doc.Revisions)
        {
            if (rev.DateTime < cutoffDate)
                revisionsToReject.Add(rev);
        }

        // Reject the collected old revisions.
        foreach (Revision rev in revisionsToReject)
        {
            rev.Reject();
        }

        // Save the resulting document.
        doc.Save("RevisionsFiltered.docx");
    }
}
