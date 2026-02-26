using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add initial content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Write("Original content. ");

        // Start tracking revisions with a specific author and timestamp.
        doc.StartTrackRevisions("Alice", DateTime.Now);

        // Insert new text – this will be recorded as an insertion revision.
        builder.Write("Inserted revision text. ");

        // Delete a run – this will be recorded as a deletion revision.
        // The first run contains "Original content. ".
        doc.FirstSection.Body.FirstParagraph.Runs[0].Remove();

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Add more text that will NOT be tracked.
        builder.Write("Non‑tracked text. ");

        // Accept all revisions – the document will contain only the final state.
        doc.AcceptAllRevisions();

        // Save the document to a DOCX file.
        string outputPath = "TrackedRevisions.docx";
        doc.Save(outputPath);
    }
}
