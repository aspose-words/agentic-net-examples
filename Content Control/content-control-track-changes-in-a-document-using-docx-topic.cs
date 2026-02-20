using System;
using Aspose.Words;
using Aspose.Words.Markup;

class TrackChangesWithContentControl
{
    static void Main()
    {
        // Create a new blank document (create rule)
        Document doc = new Document();

        // Start tracking revisions (track changes)
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Insert some text before the content control
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Paragraph before the content control.");

        // Insert a rich‑text content control (StructuredDocumentTag) into the document body
        StructuredDocumentTag richControl = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);
        richControl.Title = "SampleControl"; // optional title
        builder.InsertNode(richControl);

        // Move the builder inside the content control and add editable text
        builder.MoveTo(richControl);
        builder.Write("Editable text inside the content control.");

        // Continue with normal document content
        builder.Writeln();
        builder.Writeln("Paragraph after the content control.");

        // Stop tracking revisions (optional, can keep tracking)
        doc.StopTrackRevisions();

        // Save the document to a DOCX file (save rule)
        doc.Save("TrackedChangesWithContentControl.docx");
    }
}
