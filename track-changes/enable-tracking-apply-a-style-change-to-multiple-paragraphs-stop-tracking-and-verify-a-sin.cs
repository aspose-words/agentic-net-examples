using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start tracking revisions with a specific author.
        doc.StartTrackRevisions("Tester", DateTime.Now);

        // Apply a style (Heading1) while tracking is enabled.
        // Each paragraph insertion will be recorded as a revision.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Paragraph 1.");
        builder.Writeln("Paragraph 2.");
        builder.Writeln("Paragraph 3.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Verify that the insertions are grouped into a single revision group.
        int groupCount = doc.Revisions.Groups.Count;
        if (groupCount != 1)
            throw new InvalidOperationException($"Expected 1 revision group, but found {groupCount}.");

        RevisionGroup group = doc.Revisions.Groups[0];
        // Output basic information about the revision group.
        Console.WriteLine($"Revision group author: {group.Author}");
        Console.WriteLine($"Revision group type: {group.RevisionType}");
        Console.WriteLine($"Revision group text: {group.Text.Trim()}");

        // Save the document to the local file system.
        doc.Save("TrackedChanges.docx");
    }
}
