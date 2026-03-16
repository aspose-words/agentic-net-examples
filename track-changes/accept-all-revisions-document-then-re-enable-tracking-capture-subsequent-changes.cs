using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the existing document.
        Document doc = new Document("Input.docx");

        // Accept all tracked changes that are currently in the document.
        doc.AcceptAllRevisions();

        // Re‑enable change tracking so that any further edits are recorded as revisions.
        doc.StartTrackRevisions("Automated", DateTime.Now);

        // Example edit after re‑enabling tracking.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("New content added after accepting revisions.");

        // Save the updated document.
        doc.Save("Output.docx");
    }
}
