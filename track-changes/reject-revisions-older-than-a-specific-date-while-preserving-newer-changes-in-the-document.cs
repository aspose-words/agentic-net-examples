using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add initial content (no revision).
        builder.Writeln("Original paragraph.");

        // First set of revisions (older date).
        DateTime oldRevisionDate = new DateTime(2023, 1, 1);
        doc.StartTrackRevisions("Author1", oldRevisionDate);
        builder.Writeln("Inserted paragraph from older revision.");
        doc.StopTrackRevisions();

        // Second set of revisions (newer date).
        DateTime newRevisionDate = new DateTime(2023, 6, 1);
        doc.StartTrackRevisions("Author2", newRevisionDate);
        builder.Writeln("Inserted paragraph from newer revision.");
        doc.StopTrackRevisions();

        // Define cutoff date: revisions older than this will be rejected.
        DateTime cutoffDate = new DateTime(2023, 3, 1);

        // Process revisions: reject those older than cutoff, accept newer ones.
        // Iterate backwards because Accept/Reject modifies the collection.
        for (int i = doc.Revisions.Count - 1; i >= 0; i--)
        {
            Revision rev = doc.Revisions[i];
            if (rev.DateTime < cutoffDate)
                rev.Reject();
            else
                rev.Accept();
        }

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
