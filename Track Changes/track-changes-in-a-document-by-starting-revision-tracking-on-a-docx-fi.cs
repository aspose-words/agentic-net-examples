using System;
using Aspose.Words;

class TrackRevisionsExample
{
    static void Main()
    {
        // Load an existing DOCX file.
        Document doc = new Document("Input.docx");

        // Start tracking revisions with author name and current date/time.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Make a change that will be recorded as a revision.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This paragraph is added as a revision.");

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Save the document with revisions to a new file.
        doc.Save("OutputWithRevisions.docx");
    }
}
