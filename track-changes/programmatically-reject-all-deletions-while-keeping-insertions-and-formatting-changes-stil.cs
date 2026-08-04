using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content that will not be a revision.
        builder.Writeln("Original paragraph.");

        // Start tracking revisions.
        doc.StartTrackRevisions("Author", DateTime.Now);

        // Insert new content – this will be recorded as an insertion revision.
        builder.Writeln("Inserted paragraph.");

        // Delete the first paragraph – this will be recorded as a deletion revision.
        Paragraph paragraphToDelete = doc.FirstSection.Body.Paragraphs[0];
        paragraphToDelete.Remove();

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Reject only the deletion revisions, leaving insertions and format changes untouched.
        for (int i = doc.Revisions.Count - 1; i >= 0; i--)
        {
            Revision revision = doc.Revisions[i];
            if (revision.RevisionType == RevisionType.Deletion)
                revision.Reject();
        }

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
