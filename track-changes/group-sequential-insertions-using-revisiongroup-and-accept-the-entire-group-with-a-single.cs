using System;
using System.IO;
using Aspose.Words;

public class RevisionGroupDemo
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content that is not tracked.
        builder.Writeln("This paragraph is not tracked.");

        // Start tracking revisions.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Insert several consecutive paragraphs – they will belong to the same revision group.
        builder.Writeln("First inserted paragraph.");
        builder.Writeln("Second inserted paragraph.");
        builder.Writeln("Third inserted paragraph.");

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Ensure that a revision group was created.
        if (doc.Revisions.Groups.Count == 0)
            throw new InvalidOperationException("No revision groups were found.");

        // Accept all revisions in the document with a single call.
        // Since the only revisions present belong to the single group, this effectively accepts the whole group.
        doc.Revisions.AcceptAll();

        // After acceptance, the group should no longer exist.
        if (doc.Revisions.Groups.Count != 0)
            throw new InvalidOperationException("Revision group was not fully accepted.");

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RevisionGroupAccept.docx");
        doc.Save(outputPath);
    }
}
