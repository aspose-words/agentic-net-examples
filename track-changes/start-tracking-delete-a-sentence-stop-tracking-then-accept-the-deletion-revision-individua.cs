using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write three separate sentences – each will be a separate Run.
        builder.Write("This is the first sentence. ");
        builder.Write("This is the second sentence. ");
        builder.Write("This is the third sentence.");

        // Save the original document (optional, just to see the starting state).
        doc.Save("Original.docx");

        // Start tracking revisions.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Delete the second sentence (the second Run in the first paragraph).
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
        if (firstParagraph.Runs.Count < 2)
            throw new InvalidOperationException("Expected at least two runs in the paragraph.");

        // This removal creates a Deletion-type revision.
        firstParagraph.Runs[1].Remove();

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Verify that a revision was created.
        if (!doc.HasRevisions || doc.Revisions.Count == 0)
            throw new InvalidOperationException("No revisions were generated.");

        // Save the document while the deletion revision is still pending.
        doc.Save("WithDeletionRevision.docx");

        // Accept the deletion revision individually.
        Revision deletionRevision = null;
        foreach (Revision rev in doc.Revisions)
        {
            if (rev.RevisionType == RevisionType.Deletion)
            {
                deletionRevision = rev;
                break;
            }
        }

        if (deletionRevision == null)
            throw new InvalidOperationException("Deletion revision not found.");

        deletionRevision.Accept();

        // After acceptance, there should be no remaining revisions.
        if (doc.HasRevisions)
            throw new InvalidOperationException("Revisions still exist after acceptance.");

        // Save the final document where the sentence has been permanently removed.
        doc.Save("Final.docx");
    }
}
