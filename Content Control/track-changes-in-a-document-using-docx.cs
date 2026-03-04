using System;
using Aspose.Words;

class TrackChangesExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for editing the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write initial content that will not be tracked.
        builder.Writeln("Original paragraph. ");

        // Start tracking revisions with a specified author name.
        doc.StartTrackRevisions("John Doe");

        // Insert text that will be recorded as an insertion revision.
        builder.Writeln("This sentence is tracked as an insertion. ");

        // Stop tracking revisions; subsequent edits will not be recorded.
        doc.StopTrackRevisions();

        // Add more content that will not be tracked.
        builder.Writeln("Another untracked paragraph. ");

        // Accept all revisions in the document.
        doc.Revisions.AcceptAll();

        // Save the document to a DOCX file.
        doc.Save("TrackedChanges.docx");
    }
}
