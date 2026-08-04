using System;
using System.IO;
using Aspose.Words;

public class TrackChangesErrorHandling
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write initial text (no revision yet).
        builder.Writeln("Original paragraph.");

        // Start tracking revisions with a specific author.
        doc.StartTrackRevisions("Alice", DateTime.Now);

        // Insert a new paragraph – this will be an insertion revision.
        builder.Writeln("Inserted paragraph.");

        // Delete the first paragraph to create a deletion revision.
        Paragraph firstParagraph = doc.FirstSection.Body.Paragraphs[0];
        firstParagraph.Remove();

        // At this point we have two revisions: one insertion and one deletion.
        // Capture the deletion revision reference (it is at index 0 after the removal).
        Revision deletionRevision = doc.Revisions[0];

        // Reject the deletion revision – the paragraph will be restored.
        deletionRevision.Reject();

        // Attempt to accept the same revision again.
        // Since the revision has already been rejected, it no longer exists in the collection.
        // This operation will throw an exception, which we handle gracefully.
        try
        {
            deletionRevision.Accept();
            Console.WriteLine("Revision accepted successfully (unexpected).");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: Attempted to accept a revision that was already rejected. Message: {ex.Message}");
        }

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Save the resulting document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TrackChangesErrorHandling.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
