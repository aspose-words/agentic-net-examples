using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the sample input and output documents.
        const string inputPath = "RevisionsSample.docx";
        const string outputPath = "RevisionsAccepted.docx";

        // -----------------------------------------------------------------
        // 1. Create a sample document that contains tracked changes.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Normal content – not a revision.
        builder.Writeln("This is the original paragraph.");

        // Start tracking revisions.
        doc.StartTrackRevisions("SampleAuthor");

        // Insertion revision.
        builder.Writeln("This paragraph was inserted while tracking.");

        // Deletion revision: add a paragraph then remove it while tracking.
        builder.Writeln("This paragraph will be deleted.");
        // The paragraph we just added is the last one in the body.
        Paragraph paraToDelete = doc.FirstSection.Body.Paragraphs[doc.FirstSection.Body.Paragraphs.Count - 1];
        paraToDelete.Remove();

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Save the document that contains revisions.
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // 2. Load the document, accept all revisions, and save the cleaned version.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        // Accept every tracked change.
        loadedDoc.AcceptAllRevisions();

        // Save the result without any revisions.
        loadedDoc.Save(outputPath);
    }
}
