using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content (these edits are NOT tracked).
        builder.Writeln("Paragraph 1 - not tracked.");
        builder.Writeln("Paragraph 2 - not tracked.");

        // Start tracking revisions.
        doc.StartTrackRevisions("Test Author", DateTime.Now);

        // Make a change that will be recorded as a revision.
        // In Aspose.Words only insertions and deletions are tracked.
        builder.Writeln("Paragraph 3 - tracked insertion.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Verify that exactly one revision group was created.
        if (doc.Revisions.Groups.Count != 1)
            throw new InvalidOperationException($"Expected 1 revision group, but found {doc.Revisions.Groups.Count}.");

        RevisionGroup group = doc.Revisions.Groups[0];
        if (group.RevisionType != RevisionType.Insertion)
            throw new InvalidOperationException($"Expected revision type Insertion, but found {group.RevisionType}.");

        // Output verification result.
        Console.WriteLine("Revision group verified:");
        Console.WriteLine($"  Author: {group.Author}");
        Console.WriteLine($"  Type: {group.RevisionType}");
        Console.WriteLine($"  Text: {group.Text.Trim()}");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "TrackedInsertion.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
