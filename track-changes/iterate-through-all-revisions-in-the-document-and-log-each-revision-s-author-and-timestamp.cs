using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial text that will NOT be a revision.
        builder.Writeln("Initial content. ");

        // Start tracking revisions with author "Alice".
        doc.StartTrackRevisions("Alice", DateTime.Now);
        builder.Writeln("First revision added by Alice. ");
        // Stop tracking to finish Alice's changes.
        doc.StopTrackRevisions();

        // Start tracking revisions with author "Bob".
        doc.StartTrackRevisions("Bob", DateTime.Now);
        builder.Writeln("Second revision added by Bob. ");

        // Create a deletion revision by removing the first run.
        // The first run contains "Initial content. ".
        doc.FirstSection.Body.FirstParagraph.Runs[0].Remove();

        // Stop tracking to finish Bob's changes.
        doc.StopTrackRevisions();

        // Save the document so the revisions are persisted.
        const string outputPath = "RevisionsDemo.docx";
        doc.Save(outputPath);

        // Iterate through all revisions and log author and timestamp.
        foreach (Revision revision in doc.Revisions)
        {
            Console.WriteLine($"Author: {revision.Author}, Timestamp: {revision.DateTime}");
        }
    }
}
