using System;
using Aspose.Words;
using Aspose.Words.Markup;

class TrackChangesWithContentControl
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start tracking revisions – all subsequent changes will be recorded.
        doc.StartTrackRevisions("John Doe");

        // Insert a plain‑text content control (structured document tag).
        StructuredDocumentTag sdt = builder.InsertStructuredDocumentTag(SdtType.PlainText);
        sdt.Title = "SampleControl";
        sdt.Tag = "SampleTag";

        // Add initial text inside the content control.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Initial text inside the content control.");
        para.AppendChild(run);
        sdt.AppendChild(para);

        // Stop tracking for now.
        doc.StopTrackRevisions();

        // Make another change while tracking is active to demonstrate a second revision.
        doc.StartTrackRevisions("John Doe");
        builder.Writeln("Additional paragraph added after the content control.");
        doc.StopTrackRevisions();

        // Save the document – the .docx will show tracked changes when opened in Word.
        doc.Save("TrackedChanges.docx");
    }
}
