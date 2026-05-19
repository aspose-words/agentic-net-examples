using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Path for the resulting document.
        const string outputPath = "RevisionsFiltered.docx";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content that is not a revision.
        builder.Writeln("Original content.");

        // ---------- Create older revisions ----------
        DateTime oldRevisionDate = new DateTime(2020, 1, 1);
        doc.StartTrackRevisions("Author1", oldRevisionDate);
        builder.Writeln("This text was added on an old date.");
        doc.StopTrackRevisions();

        // ---------- Create newer revisions ----------
        DateTime newRevisionDate = new DateTime(2023, 1, 1);
        doc.StartTrackRevisions("Author2", newRevisionDate);
        builder.Writeln("This text was added on a newer date.");
        doc.StopTrackRevisions();

        // Define the cutoff date: revisions older than this will be rejected.
        DateTime cutoffDate = new DateTime(2021, 1, 1);

        // Iterate backwards through the revision collection because rejecting a revision
        // modifies the collection.
        for (int i = doc.Revisions.Count - 1; i >= 0; i--)
        {
            Revision rev = doc.Revisions[i];
            if (rev.DateTime < cutoffDate)
            {
                rev.Reject(); // Reject the old revision.
            }
        }

        // Save the document with only the newer revisions preserved.
        doc.Save(outputPath);
    }
}
