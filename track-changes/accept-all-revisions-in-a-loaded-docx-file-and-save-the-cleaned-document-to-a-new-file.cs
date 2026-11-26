using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample document and generate some revisions.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Initial content (no revisions yet).
        builder.Writeln("This is the original paragraph.");

        // Start tracking revisions.
        doc.StartTrackRevisions("SampleAuthor", DateTime.Now);

        // Insert a new paragraph (will be an insertion revision).
        builder.Writeln("This paragraph was inserted while tracking.");

        // Insert another paragraph at the beginning of the document.
        builder.MoveToDocumentStart();
        builder.Writeln("Inserted at the beginning of the document.");

        // Delete the original paragraph (will be a deletion revision).
        // After the two insertions the original paragraph is at index 2.
        Paragraph originalParagraph = doc.FirstSection.Body.Paragraphs[2];
        originalParagraph.Remove();

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Save the document that contains revisions.
        const string originalPath = "OriginalWithRevisions.docx";
        doc.Save(originalPath);

        // -----------------------------------------------------------------
        // 2. Load the document, accept all revisions, and save the cleaned file.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(originalPath);

        // Accept every revision in the document.
        loadedDoc.AcceptAllRevisions();

        // Save the cleaned document (no remaining revisions).
        const string cleanedPath = "CleanedDocument.docx";
        loadedDoc.Save(cleanedPath);
    }
}
