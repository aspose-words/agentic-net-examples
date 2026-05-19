using System;
using Aspose.Words;

public class TrackChangesDemo
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable tracking before making any changes.
        doc.StartTrackRevisions("DemoAuthor", DateTime.Now);

        // Add three paragraphs while tracking is enabled.
        // All insertions made in this block will be recorded as revisions.
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");
        builder.Writeln("Third paragraph.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Verify that revisions were created.
        if (!doc.HasRevisions)
            throw new InvalidOperationException("No revisions were recorded.");

        // Verify that there is exactly one revision group (the three insertions are sequential and form a single group).
        if (doc.Revisions.Groups.Count != 1)
            throw new InvalidOperationException($"Expected 1 revision group, but found {doc.Revisions.Groups.Count}.");

        // Inspect the revision group details.
        RevisionGroup group = doc.Revisions.Groups[0];
        Console.WriteLine($"Revision group author: {group.Author}");
        Console.WriteLine($"Revision group type: {group.RevisionType}");
        Console.WriteLine($"Revision group text: {group.Text.Trim()}");

        // Save the document to verify the changes.
        doc.Save("TrackChangesDemo.docx");
    }
}
