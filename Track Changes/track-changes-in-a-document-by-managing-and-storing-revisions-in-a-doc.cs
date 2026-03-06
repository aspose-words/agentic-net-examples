using System;
using Aspose.Words;

class RevisionDemo
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add initial content that will NOT be tracked.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Write("Initial content. ");

        // Start tracking revisions. All subsequent changes will be recorded.
        // Author name and current date/time are stored with each revision.
        doc.StartTrackRevisions("Alice", DateTime.Now);

        // Insert text that will appear as an insertion revision.
        builder.Write("First revision. ");
        builder.Writeln("Second revision line.");

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Create a deletion revision by removing the first run (the word "Initial").
        doc.FirstSection.Body.FirstParagraph.Runs[0].Remove();

        // Optional: accept all revisions so the document reflects the final state.
        // doc.AcceptAllRevisions();

        // Save the document with its tracked changes to a DOCX file.
        doc.Save("TrackedChanges.docx");
    }
}
