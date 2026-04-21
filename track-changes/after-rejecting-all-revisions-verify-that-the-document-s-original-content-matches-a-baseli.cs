using System;
using Aspose.Words;

public class TrackChangesDemo
{
    public static void Main()
    {
        // Define file names.
        const string baselinePath = "Baseline.docx";
        const string editedPath = "EditedWithRevisions.docx";
        const string finalPath = "FinalAfterReject.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a baseline document with original content.
        // -----------------------------------------------------------------
        Document baselineDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(baselineDoc);
        builder.Writeln("This is the original paragraph.");
        builder.Writeln("Second line of original content.");
        baselineDoc.Save(baselinePath);

        // -----------------------------------------------------------------
        // Step 2: Load the baseline, start tracking revisions, and modify it.
        // -----------------------------------------------------------------
        Document docWithRevisions = new Document(baselinePath);
        docWithRevisions.StartTrackRevisions("Author", DateTime.Now);

        // Insert a new paragraph (insertion revision).
        DocumentBuilder revBuilder = new DocumentBuilder(docWithRevisions);
        revBuilder.Writeln("This paragraph was inserted as a revision.");

        // Delete the first paragraph (deletion revision).
        Paragraph firstParagraph = docWithRevisions.FirstSection.Body.Paragraphs[0];
        firstParagraph.Remove();

        docWithRevisions.StopTrackRevisions();
        docWithRevisions.Save(editedPath);

        // -----------------------------------------------------------------
        // Step 3: Reject all revisions, reverting the document to its original state.
        // -----------------------------------------------------------------
        // Use the RevisionCollection.RejectAll method (Document.RejectAllRevisions does not exist).
        docWithRevisions.Revisions.RejectAll();

        // Save the final document after rejection.
        docWithRevisions.Save(finalPath);

        // -----------------------------------------------------------------
        // Step 4: Verify that the content matches the baseline document.
        // -----------------------------------------------------------------
        string baselineText = baselineDoc.GetText().Trim();
        string finalText = docWithRevisions.GetText().Trim();

        if (!baselineText.Equals(finalText, StringComparison.Ordinal))
        {
            throw new InvalidOperationException("The document content after rejecting revisions does not match the baseline.");
        }

        Console.WriteLine("Revisions rejected successfully; document matches the baseline.");
    }
}
