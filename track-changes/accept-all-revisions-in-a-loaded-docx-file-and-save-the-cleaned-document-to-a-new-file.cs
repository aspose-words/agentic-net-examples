using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the sample and cleaned documents.
        string sourcePath = "sample_with_revisions.docx";
        string cleanedPath = "cleaned.docx";

        // -----------------------------------------------------------------
        // 1. Create a sample document and generate some revisions.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Initial content (no revisions).
        builder.Writeln("This is the original paragraph.");

        // Enable tracking to create revisions.
        doc.StartTrackRevisions("SampleAuthor", DateTime.Now);

        // Insert new text – this will be recorded as an insertion revision.
        builder.Writeln("This paragraph was added while tracking revisions.");

        // Delete a paragraph – also recorded as a deletion revision.
        builder.MoveToDocumentStart();
        builder.CurrentParagraph.Remove();

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Save the document that contains revisions.
        doc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the document with revisions.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // -----------------------------------------------------------------
        // 3. Accept all revisions, removing revision marks.
        // -----------------------------------------------------------------
        loadedDoc.AcceptAllRevisions();

        // -----------------------------------------------------------------
        // 4. Save the cleaned document to a new file.
        // -----------------------------------------------------------------
        loadedDoc.Save(cleanedPath);
    }
}
