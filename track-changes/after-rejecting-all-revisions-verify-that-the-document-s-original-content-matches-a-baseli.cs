using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // File names for the baseline document and the document that will contain revisions.
        const string baselinePath = "Baseline.docx";
        const string revisedPath = "Revised.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a baseline document that contains the original text.
        // -----------------------------------------------------------------
        Document baselineDoc = new Document();
        DocumentBuilder baselineBuilder = new DocumentBuilder(baselineDoc);
        baselineBuilder.Writeln("Original line 1.");
        baselineBuilder.Writeln("Original line 2.");
        baselineDoc.Save(baselinePath);

        // -----------------------------------------------------------------
        // Step 2: Load the baseline document and start tracking revisions.
        // -----------------------------------------------------------------
        Document doc = new Document(baselinePath);
        doc.StartTrackRevisions("TestAuthor", DateTime.Now);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the end of the document so the inserted paragraph
        // is added after the original content (otherwise it would be removed again).
        builder.MoveToDocumentEnd();
        builder.Writeln("Inserted line.");

        // Delete the first original paragraph – this will be recorded as a deletion revision.
        Paragraph firstParagraph = doc.FirstSection.Body.Paragraphs[0];
        firstParagraph.Remove();

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Ensure that revisions were actually created.
        if (!doc.HasRevisions || doc.Revisions.Count == 0)
            throw new InvalidOperationException("No revisions were generated.");

        // -----------------------------------------------------------------
        // Step 3: Reject all revisions, reverting the document to its original state.
        // -----------------------------------------------------------------
        doc.Revisions.RejectAll();

        // Save the document after rejecting revisions.
        doc.Save(revisedPath);

        // -----------------------------------------------------------------
        // Step 4: Verify that the content matches the baseline.
        // -----------------------------------------------------------------
        Document baselineCheck = new Document(baselinePath);
        string originalText = baselineCheck.GetText();
        string finalText = doc.GetText();

        if (!string.Equals(originalText, finalText, StringComparison.Ordinal))
            throw new Exception("The document content after rejecting revisions does not match the baseline.");

        Console.WriteLine("Revisions rejected successfully; document matches the baseline.");
    }
}
