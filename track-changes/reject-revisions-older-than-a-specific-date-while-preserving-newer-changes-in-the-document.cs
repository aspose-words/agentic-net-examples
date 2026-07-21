using System;
using Aspose.Words;

public class RejectOldRevisions
{
    public static void Main()
    {
        // Define the cutoff date. Revisions older than this will be rejected.
        DateTime cutoffDate = new DateTime(2023, 1, 1);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial content that is not a revision.
        builder.Writeln("Original content. ");

        // ----- Create an old revision (date before the cutoff) -----
        doc.StartTrackRevisions("OldAuthor", new DateTime(2022, 12, 31));
        builder.Writeln("This text is an old revision.");
        doc.StopTrackRevisions();

        // ----- Create a new revision (date after the cutoff) -----
        doc.StartTrackRevisions("NewAuthor", new DateTime(2023, 6, 1));
        builder.Writeln("This text is a new revision.");
        doc.StopTrackRevisions();

        // Iterate over the revisions in reverse order to safely modify the collection.
        for (int i = doc.Revisions.Count - 1; i >= 0; i--)
        {
            Revision rev = doc.Revisions[i];
            if (rev.DateTime < cutoffDate)
            {
                // Reject revisions older than the cutoff date.
                rev.Reject();
            }
            // Newer revisions are left untouched (preserved).
        }

        // Save the resulting document.
        doc.Save("RejectedOldRevisions.docx");
    }
}
