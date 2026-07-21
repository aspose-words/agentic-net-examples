using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the sample files.
        const string originalPath = "OriginalWithRevisions.docx";
        const string cleanedPath = "Cleaned.docx";

        // -----------------------------------------------------------------
        // 1. Create a sample document that contains tracked changes.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable tracking so that subsequent edits become revisions.
        doc.StartTrackRevisions("Sample Author");

        // Add some text – each write creates an insertion revision.
        builder.Writeln("First line with revision.");
        builder.Writeln("Second line with revision.");

        // Stop tracking to prevent further edits from being marked.
        doc.StopTrackRevisions();

        // Save the document that now holds revisions.
        doc.Save(originalPath);

        // -----------------------------------------------------------------
        // 2. Load the document that contains revisions.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(originalPath);

        // -----------------------------------------------------------------
        // 3. Accept all revisions, removing the change markup.
        // -----------------------------------------------------------------
        loadedDoc.AcceptAllRevisions();

        // Optional sanity check – the document should have no revisions left.
        if (loadedDoc.HasRevisions)
            throw new InvalidOperationException("Revisions were not fully accepted.");

        // -----------------------------------------------------------------
        // 4. Save the cleaned document to a new file.
        // -----------------------------------------------------------------
        loadedDoc.Save(cleanedPath);
    }
}
