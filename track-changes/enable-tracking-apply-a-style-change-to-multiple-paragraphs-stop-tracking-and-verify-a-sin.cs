using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add several paragraphs with sample text.
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");
        builder.Writeln("Third paragraph.");

        // Start tracking revisions with a specific author.
        doc.StartTrackRevisions("Test Author", DateTime.Now);

        // Apply a style change to each paragraph while tracking is enabled.
        foreach (Paragraph para in doc.FirstSection.Body.Paragraphs)
        {
            para.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        }

        // Insert a dummy run to generate a revision (format changes are not tracked as revisions).
        // This will create a single insertion revision that groups the changes.
        builder.Writeln("Dummy text to create a revision.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Save the document.
        doc.Save("TrackedChanges.docx");

        // Verify that a single revision group was created.
        if (doc.Revisions.Groups.Count != 1)
            throw new Exception($"Expected 1 revision group, but found {doc.Revisions.Groups.Count}.");

        RevisionGroup group = doc.Revisions.Groups[0];

        // The revision type for an insertion is RevisionType.Insertion.
        if (group.RevisionType != RevisionType.Insertion)
            throw new Exception($"Expected revision type Insertion, but found {group.RevisionType}.");

        // Confirm the author of the revision group.
        if (group.Author != "Test Author")
            throw new Exception($"Expected revision author 'Test Author', but found '{group.Author}'.");

        Console.WriteLine("Revision tracking and verification completed successfully.");
    }
}
