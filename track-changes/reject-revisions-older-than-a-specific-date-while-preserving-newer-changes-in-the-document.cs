using System;
using System.Collections.Generic;
using Aspose.Words;

public class Program
{
    public static void Main(string[] args)
    {
        // Create a new document and add initial content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Original paragraph.");

        // First set of revisions (older date).
        DateTime oldRevisionDate = new DateTime(2020, 1, 1);
        doc.StartTrackRevisions("OldAuthor", oldRevisionDate);
        builder.Writeln("Old revision paragraph.");
        doc.StopTrackRevisions();

        // Second set of revisions (newer date).
        DateTime newRevisionDate = new DateTime(2022, 1, 1);
        doc.StartTrackRevisions("NewAuthor", newRevisionDate);
        builder.Writeln("New revision paragraph.");
        doc.StopTrackRevisions();

        // Define the cutoff date: revisions older than this will be rejected.
        DateTime cutoffDate = new DateTime(2021, 1, 1);

        // Collect revisions to reject (those older than the cutoff date).
        List<Revision> revisionsToReject = new List<Revision>();
        foreach (Revision rev in doc.Revisions)
        {
            if (rev.DateTime < cutoffDate)
                revisionsToReject.Add(rev);
        }

        // Reject the collected revisions.
        foreach (Revision rev in revisionsToReject)
        {
            rev.Reject();
        }

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
