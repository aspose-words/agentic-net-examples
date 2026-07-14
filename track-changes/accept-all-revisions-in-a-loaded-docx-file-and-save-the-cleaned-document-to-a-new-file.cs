using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary files.
        string originalPath = "OriginalWithRevisions.docx";
        string cleanedPath = "Cleaned.docx";

        // -----------------------------------------------------------------
        // 1. Create a sample document and generate some revisions.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Normal text – not a revision.
        builder.Writeln("This text will not be a revision.");

        // Start tracking revisions.
        doc.StartTrackRevisions("Author", DateTime.Now);

        // These writes will be recorded as revisions.
        builder.Writeln("First revision line.");
        builder.Writeln("Second revision line.");

        // Stop tracking – further edits are not revisions.
        doc.StopTrackRevisions();

        // Save the document that contains revisions.
        doc.Save(originalPath);

        // -----------------------------------------------------------------
        // 2. Load the document with revisions, accept all of them, and save.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(originalPath);

        // Ensure the document actually has revisions before proceeding.
        if (!loadedDoc.HasRevisions)
            throw new InvalidOperationException("The loaded document does not contain any revisions.");

        // Accept every revision in the document.
        loadedDoc.AcceptAllRevisions();

        // Save the cleaned document (no revisions remain).
        loadedDoc.Save(cleanedPath);
    }
}
