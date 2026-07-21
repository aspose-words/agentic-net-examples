using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial content that will NOT be a revision.
        builder.Writeln("Original paragraph. ");

        // Start tracking revisions. All subsequent changes will be recorded.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Insert several paragraphs sequentially. These insertions will be grouped together.
        builder.Writeln("First inserted paragraph.");
        builder.Writeln("Second inserted paragraph.");
        builder.Writeln("Third inserted paragraph.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // At this point the document contains a single revision group for the sequential insertions.
        Console.WriteLine($"Revision groups count: {doc.Revisions.Groups.Count}");
        if (doc.Revisions.Groups.Count > 0)
        {
            RevisionGroup group = doc.Revisions.Groups[0];
            Console.WriteLine($"Group author: {group.Author}");
            Console.WriteLine($"Group type: {group.RevisionType}");
            Console.WriteLine($"Group text: {group.Text.Trim()}");
        }

        // Accept the entire group of revisions with a single call.
        // Since the group is the only set of revisions, AcceptAll() suffices.
        doc.Revisions.AcceptAll();

        // Verify that revisions have been accepted.
        Console.WriteLine($"Revisions after accept: {doc.Revisions.Count}");

        // Save the resulting document.
        doc.Save("RevisionGroupAccepted.docx");
    }
}
