using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content that is not tracked.
        builder.Writeln("Original paragraph.");

        // Start tracking revisions with a specific author.
        doc.StartTrackRevisions("Alice", DateTime.Now);

        // Insert several paragraphs sequentially. These insertions will be grouped together.
        builder.Writeln("Inserted paragraph 1.");
        builder.Writeln("Inserted paragraph 2.");
        builder.Writeln("Inserted paragraph 3.");

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Verify that a revision group was created for the sequential insertions.
        if (doc.Revisions.Groups.Count != 1)
            throw new InvalidOperationException("Expected exactly one revision group for the insertions.");

        // Accept the entire group of revisions with a single call.
        // Since the group is the only set of revisions, AcceptAll will accept them all.
        doc.Revisions.AcceptAll();

        // Optionally, verify that all revisions have been accepted.
        if (doc.HasRevisions)
            throw new InvalidOperationException("Revisions were not fully accepted.");

        // Save the resulting document.
        doc.Save("RevisionGroupExample.docx");
    }
}
