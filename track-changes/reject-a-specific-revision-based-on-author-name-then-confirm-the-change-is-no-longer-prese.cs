using System;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add initial content.
        builder.Writeln("Original text.");

        // Start tracking revisions with a specific author.
        string targetAuthor = "John Doe";
        doc.StartTrackRevisions(targetAuthor, DateTime.Now);

        // Make a change that will be recorded as a revision.
        builder.Writeln("Inserted text.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Collect revisions authored by the target author, then reject them.
        var revisionsToReject = doc.Revisions
                                   .Cast<Revision>()
                                   .Where(r => r.Author == targetAuthor)
                                   .ToList();

        foreach (Revision rev in revisionsToReject)
        {
            rev.Reject();
        }

        // Verify that the inserted text has been removed.
        if (doc.GetText().Contains("Inserted text"))
            throw new Exception("The revision was not rejected correctly.");

        // Save the resulting document.
        doc.Save("result.docx");
    }
}
