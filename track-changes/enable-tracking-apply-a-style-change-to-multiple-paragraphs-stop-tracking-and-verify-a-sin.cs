using System;
using Aspose.Words;

public class TrackChangesExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add three normal paragraphs (these edits are not tracked).
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");
        builder.Writeln("Third paragraph.");

        // Start tracking revisions with a specific author.
        doc.StartTrackRevisions("Sample Author", DateTime.Now);

        // Insert a new paragraph while tracking is enabled.
        // This insertion will be recorded as a revision.
        builder.Writeln("Inserted paragraph with Heading 1 style.");
        // Apply a style change to the newly inserted paragraph.
        // The style change itself is not tracked as a revision, but the insertion is.
        builder.CurrentParagraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Verify that the insertion revision is grouped into a single revision group.
        int groupCount = doc.Revisions.Groups.Count;
        if (groupCount != 1)
        {
            throw new InvalidOperationException($"Expected 1 revision group, but found {groupCount}.");
        }

        // Inspect the revision group details (author, type, text).
        RevisionGroup group = doc.Revisions.Groups[0];
        Console.WriteLine($"Revision group author: {group.Author}");
        Console.WriteLine($"Revision group type: {group.RevisionType}");
        Console.WriteLine($"Revision group text: {group.Text.Trim()}");

        // Save the document to verify the revisions visually if needed.
        doc.Save("TrackedChanges.docx");
    }
}
