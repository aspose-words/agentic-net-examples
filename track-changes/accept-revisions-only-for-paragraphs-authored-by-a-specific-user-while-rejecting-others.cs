using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder for editing.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Author1 makes a revision.
        doc.StartTrackRevisions("Author1", DateTime.Now);
        builder.Writeln("Paragraph added by Author1.");
        doc.StopTrackRevisions();

        // Author2 makes a revision.
        doc.StartTrackRevisions("Author2", DateTime.Now);
        builder.Writeln("Paragraph added by Author2.");
        doc.StopTrackRevisions();

        // Create a snapshot of the revisions to avoid modifying the collection while iterating.
        List<Revision> revisions = doc.Revisions.Cast<Revision>().ToList();

        // Accept revisions authored by Author1, reject all others.
        foreach (Revision rev in revisions)
        {
            if (rev.Author == "Author1")
                rev.Accept();
            else
                rev.Reject();
        }

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
