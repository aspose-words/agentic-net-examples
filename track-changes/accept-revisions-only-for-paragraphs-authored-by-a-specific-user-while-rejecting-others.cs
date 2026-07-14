using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial content that is not a revision.
        builder.Writeln("Original paragraph. ");

        // ---------- Revisions by Alice ----------
        doc.StartTrackRevisions("Alice", DateTime.Now);
        builder.Writeln("Alice's inserted paragraph. ");
        // Delete a run to create a deletion revision for Alice.
        doc.FirstSection.Body.FirstParagraph.Runs[0].Remove();
        doc.StopTrackRevisions();

        // ---------- Revisions by Bob ----------
        doc.StartTrackRevisions("Bob", DateTime.Now);
        builder.Writeln("Bob's inserted paragraph. ");
        // Delete a run to create a deletion revision for Bob.
        doc.FirstSection.Body.FirstParagraph.Runs[0].Remove();
        doc.StopTrackRevisions();

        // At this point the document contains revisions from both authors.
        // Accept only revisions authored by Alice and reject all others.
        string targetAuthor = "Alice";

        // Copy revisions to a list because the collection changes during accept/reject.
        List<Revision> revisions = doc.Revisions.Cast<Revision>().ToList();

        foreach (Revision rev in revisions)
        {
            if (rev.Author == targetAuthor)
                rev.Accept();
            else
                rev.Reject();
        }

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
