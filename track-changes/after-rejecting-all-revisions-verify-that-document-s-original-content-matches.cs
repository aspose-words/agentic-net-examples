using System;
using Aspose.Words;

class VerifyDocumentAfterRejectingRevisions
{
    static void Main()
    {
        // Create the baseline document that represents the expected original content.
        Document baselineDoc = new Document();
        DocumentBuilder baselineBuilder = new DocumentBuilder(baselineDoc);
        baselineBuilder.Writeln("Hello world!");

        // Clone the baseline to create a document that will contain revisions.
        Document docWithRevisions = (Document)baselineDoc.Clone();
        docWithRevisions.TrackRevisions = true;

        // Make a change that will be recorded as a revision.
        DocumentBuilder revBuilder = new DocumentBuilder(docWithRevisions);
        revBuilder.Writeln("This line is a revision.");

        // Reject all revisions so the document reverts to its original state.
        docWithRevisions.Revisions.RejectAll();

        // Compare the textual content of both documents.
        bool contentsMatch = docWithRevisions.GetText() == baselineDoc.GetText();

        // Output the verification result.
        Console.WriteLine(contentsMatch
            ? "The document matches the baseline after rejecting revisions."
            : "The document does NOT match the baseline after rejecting revisions.");
    }
}
