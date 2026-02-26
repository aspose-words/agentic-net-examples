using System;
using Aspose.Words;

class RevisionTrackingExample
{
    static void Main()
    {
        // Load an existing DOCX file.
        Document doc = new Document("Input.docx");

        // Begin tracking revisions. Provide author name and timestamp.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Perform some edits that will be recorded as revisions.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This paragraph was added while tracking revisions.");

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Save the document; the inserted text will appear as a revision in Word.
        doc.Save("OutputWithRevisions.docx");
    }
}
