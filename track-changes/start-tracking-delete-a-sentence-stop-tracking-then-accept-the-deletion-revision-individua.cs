using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add three separate runs (sentences) to the first paragraph.
        builder.Write("First sentence. ");
        builder.Write("Second sentence. ");
        builder.Write("Third sentence.");
        builder.Writeln(); // End the paragraph.

        // Start tracking revisions.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Delete the second sentence (the second run) to generate a deletion revision.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;
        if (paragraph.Runs.Count < 2)
            throw new InvalidOperationException("Expected at least two runs in the paragraph.");

        // Remove the run that contains the second sentence.
        paragraph.Runs[1].Remove();

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Verify that a revision was created.
        if (!doc.HasRevisions || doc.Revisions.Count == 0)
            throw new InvalidOperationException("No revisions were generated.");

        // Find the deletion revision and accept it individually.
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

        // Accept only this deletion revision.
        deletionRevision.Accept();

        // At this point the revision collection should be empty.
        if (doc.HasRevisions)
            throw new InvalidOperationException("Revisions were not fully accepted.");

        // Save the resulting document.
        doc.Save("TrackChangesDemo.docx");
    }
}
