using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define a cutoff date. Revisions older than this will be rejected.
        DateTime cutoffDate = new DateTime(2023, 1, 1);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content that is not a revision.
        builder.Writeln("Initial content. ");

        // ----- Create an older revision (date before the cutoff) -----
        doc.StartTrackRevisions("Author", new DateTime(2022, 12, 1));
        builder.Writeln("This is an old revision that should be rejected.");
        doc.StopTrackRevisions();

        // ----- Create a newer revision (date after the cutoff) -----
        doc.StartTrackRevisions("Author", new DateTime(2023, 2, 1));
        builder.Writeln("This is a new revision that should be kept.");
        doc.StopTrackRevisions();

        // Iterate over a snapshot of the revisions collection.
        // Reject revisions older than the cutoff date, accept newer ones.
        var revisions = doc.Revisions.Cast<Revision>().ToList();
        foreach (Revision rev in revisions)
        {
            if (rev.DateTime < cutoffDate)
                rev.Reject();   // Remove old revision.
            else
                rev.Accept();   // Preserve new revision.
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
        doc.Save(outputPath);
    }
}
