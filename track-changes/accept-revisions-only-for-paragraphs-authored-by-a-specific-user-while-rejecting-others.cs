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

        // Add a paragraph that is not a revision (normal editing).
        builder.Writeln("Original paragraph.");

        // Start tracking revisions for author "Alice".
        doc.StartTrackRevisions("Alice", DateTime.Now);
        builder.Writeln("Paragraph added by Alice.");
        doc.StopTrackRevisions();

        // Start tracking revisions for author "Bob".
        doc.StartTrackRevisions("Bob", DateTime.Now);
        builder.Writeln("Paragraph added by Bob.");
        doc.StopTrackRevisions();

        // Accept only revisions authored by "Alice" and reject all others.
        // Copy the revisions to a list to avoid modifying the collection while iterating.
        List<Revision> revisions = new List<Revision>(doc.Revisions);
        foreach (Revision rev in revisions)
        {
            if (rev.Author == "Alice")
                rev.Accept();
            else
                rev.Reject();
        }

        // Save the processed document.
        doc.Save("Result.docx");
    }
}
